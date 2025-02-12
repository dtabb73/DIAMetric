using System;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Globalization;
using System.Threading;

namespace DIAMetric
{
    static class Program
    {
        static void Main(string[] args)
        {
            // Use periods to separate decimals
            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
            Console.WriteLine("DIAMetric: Quality metrics for Data-Independent Acquisition experiments");
            Console.WriteLine("David L. Tabb, ERIBA of University Medical Center of Groningen");
            Console.WriteLine("alpha version 20250212");
	    //Console.WriteLine("--DIANN: read report.pr_matrix.tsv peptide matrix");

	    var ReadDIANN = false;
	    foreach (var item in args)
	    {
		switch(item)
		{
		    case "--DIANN":
			ReadDIANN = true;
			break;
		    default:
			Console.Error.WriteLine("\tError: I don't understand this argument: {0}.", item);
			break;
		}
	    }
            var CWD = Directory.GetCurrentDirectory();
            const string mzMLPattern = "*.mzML";
            var mzMLs = Directory.GetFiles(CWD, mzMLPattern);
            var Raws = new LCMSMSExperiment();
            var RawsRunner = Raws;
            string Basename;
	    Stopwatch Timer = new Stopwatch();
	    TimeSpan Duration;
	    Timer.Start();
            Console.WriteLine("\nImporting from mzML files...");
            foreach (var current in mzMLs)
            {
                Basename = Path.GetFileNameWithoutExtension(current);
                Console.WriteLine("\tReading mzML {0}", Basename);
                RawsRunner.Next = new LCMSMSExperiment();
                var FileSpec = Path.Combine(CWD, current);
                var XMLfile = XmlReader.Create(FileSpec);
                RawsRunner = RawsRunner.Next;
                RawsRunner.SourceFile = Basename;
                RawsRunner.ReadFromMZML(XMLfile);
                RawsRunner.ParseScanNumbers();
            }
	    Timer.Stop();
	    Duration = Timer.Elapsed;
	    Console.WriteLine("\tTime for mzML reading: {0}",Duration.ToString());
	    if (ReadDIANN == true)
	    {
		Timer.Reset();
		Timer.Start();
		Console.WriteLine("\nReading DIANN result is not implemented yet.");
		Timer.Stop();
		Duration = Timer.Elapsed;
		Console.WriteLine("\tTime for DIANN reading: {0}",Duration.ToString());
	    }
	    Timer.Reset();
	    Timer.Start();
            Console.WriteLine("\nComputing metrics for each isolation window of each experiment...");
	    RawsRunner = Raws.Next;
	    while (RawsRunner != null) {
		RawsRunner.ComputeMetricsForSwaths();
		RawsRunner = RawsRunner.Next;
	    }
	    Timer.Stop();
	    Duration = Timer.Elapsed;
	    Console.WriteLine("\tTime for isolation window evaluation: {0}",Duration.ToString());

            Console.WriteLine("\nWriting DIAMetric TSV and mzQC reports...");
            Raws.WriteTextQCReport();
	    Timer.Stop();
	    Duration = Timer.Elapsed;
	    Console.WriteLine("\tTime for reporting: {0}",Duration.ToString());
        }

    }

    class ScanMetrics
    {
        public string NativeID = "";
        public int ScanNumber;
        public float ScanStartTime;
        public int mzMLPeakCount;
	public float FAIMScompVoltage=0f;
	public float mzMLtic;
	public double IsolationTarget;
	public double IsolationLowerOffset;
	public double IsolationHigherOffset;
        public ScanMetrics Next;
        //TODO Allow users to specify alternative mass for Cys
        // Values from https://education.expasy.org/student_projects/isotopident/htdocs/aa-list.html
        public static double[] AminoAcids = {57.02146,71.03711,87.03203,97.05276,99.06841,101.04768,103.00919,113.08406,114.04293,
                          115.02694,128.05858,128.09496,129.04259,131.04049,137.05891,147.06841,156.10111,
                          163.06333,186.07931};
        public static string[] AminoAcidSymbols = { "G", "A", "S", "P", "V", "T", "C", "L/I", "N",
						    "D", "Q", "K", "E", "M", "H", "F", "R", "Y", "W" };

	public ScanMetrics ExciseAllAtFirstIsolationTarget()
	{
	    // Separate all scans that match the first isolation window in the current list.  Return that sublist.
	    ScanMetrics SMRunner = this.Next;
	    if (SMRunner == null) return null;
	    double TargetMZ = SMRunner.IsolationTarget;
	    double TargetFAIMS = SMRunner.FAIMScompVoltage;
	    // Console.WriteLine("Extracting an isolation window at " + TargetMZ.ToString() + " with FAIMS CV of " + TargetFAIMS.ToString());
	    int HowMany = 0;
	    ScanMetrics SMHeader = new ScanMetrics();
	    ScanMetrics SMRunner2 = SMHeader;
	    SMRunner = this;
	    while (SMRunner != null)
	    {
		if (SMRunner.Next != null)
		{
		    if ( (SMRunner.Next.FAIMScompVoltage == TargetFAIMS) && (SMRunner.Next.IsolationTarget == TargetMZ) )
		    {
			HowMany++;
			SMRunner2.Next = SMRunner.Next;
			SMRunner2 = SMRunner2.Next;
			SMRunner.Next = SMRunner2.Next;
			SMRunner2.Next = null;
		    }
		    else SMRunner = SMRunner.Next;
		}
		else SMRunner = SMRunner.Next;
	    }
	    // Console.WriteLine("\t" + HowMany.ToString() + " MS/MS measurements collected.");
	    return SMHeader;
	}
	
	public string Stringify()
	{
	    return (this.FAIMScompVoltage.ToString() + " " + this.IsolationTarget.ToString() + " " + this.ScanStartTime.ToString());
	}

	public bool ComesBefore(ScanMetrics Other)
	{
	    if (this.FAIMScompVoltage < Other.FAIMScompVoltage)
	    {
		return true;
	    }
	    if (this.IsolationTarget < Other.IsolationTarget)
	    {
		return true;
	    }
	    if (this.ScanStartTime < Other.ScanStartTime)
	    {
		return true;
	    }
	    return false;
	}
    }

    class SWATHMetrics
    {
	public double LoMZ = 0;
	public double HiMZ = 0;
	public double WidthMZ = 0;
	public float FAIMS = 0;
	public int   MSMSCount = 0;
	public float LoRT = 0;
	public float HiRT = 0;
	public float CycleTimeMedian = 0;
	public float TIC25ileRT = 0;
	public float TIC50ileRT = 0;
	public float TIC75ileRT = 0;
	public float TotalTIC = 0;
	public float PkCount25ile = 0;
	public float PkCount50ile = 0;
	public float PkCount75ile = 0;
	public float PkCountMax = 0;
	public SWATHMetrics Next;
    }
    
    class LCMSMSExperiment
    {
        // Fields read directly from file
        public string SourceFile = "";
        public string Instrument = "";
        public string SerialNumber = "";
        public string StartTimeStamp = "";
        public float MaxScanStartTime;
        // Computed fields
        public int mzMLMS1Count;
        public int mzMLMSnCount;
	public int SWATHCount = 0;
	public int SWATHCycleCountMin = Int32.MaxValue;
	public int SWATHCycleCountMax = 0;
	public double LoMZRange = Double.PositiveInfinity;
	public double HiMZRange = 0f;
	public double SWATHWidest = 0f;
	public double SWATHNarrowest = Double.PositiveInfinity;
	public float MedianCycleTimeAverage = 0f;
	public float TIC50ileRTMin = Single.PositiveInfinity;
	public float TIC50ileRTMax = 0f;
	public float TotalTICMin = Single.PositiveInfinity;
	public float TotalTICMax = 0f;
	public float PkCount50ileMin = Single.PositiveInfinity;
	public float PkCount50ileMax = 0f;
        public static int MaxPkCount = 10000;
	public int[] mzMLPeakCountDistn = new int[MaxPkCount + 1];
        // Per-scan metrics
        public ScanMetrics ScansTable = new ScanMetrics();
	public SWATHMetrics SWATHs = new SWATHMetrics();
        private ScanMetrics ScansRunner;
        public LCMSMSExperiment Next;
        private int LastPeakCount;
        private string LastNativeID = "";

	public void ComputeMetricsForSwaths()
	{
	    ScanMetrics  TheseScans;
	    ScanMetrics  SMRunner;
	    SWATHMetrics SWATHRunner = this.SWATHs;
	    TheseScans = this.ScansTable.ExciseAllAtFirstIsolationTarget();
	    float SumMedianCycleTime = 0;
	    while (TheseScans != null)
	    {
		//Compute SWATH metrics
		SWATHRunner.Next = new SWATHMetrics();
		SWATHRunner = SWATHRunner.Next;
		this.SWATHCount++;
		// We live dangerously, assuming ScanMetrics are ordered in retention time
		SMRunner = TheseScans.Next;
		SWATHRunner.LoRT = SMRunner.ScanStartTime;
		SWATHRunner.LoMZ = SMRunner.IsolationTarget - SMRunner.IsolationLowerOffset;
		SWATHRunner.HiMZ = SMRunner.IsolationTarget + SMRunner.IsolationHigherOffset;
		SWATHRunner.WidthMZ = SWATHRunner.HiMZ - SWATHRunner.LoMZ;
		SWATHRunner.FAIMS = SMRunner.FAIMScompVoltage;
		if (SWATHRunner.LoMZ < this.LoMZRange) this.LoMZRange = SWATHRunner.LoMZ;
		if (SWATHRunner.HiMZ > this.HiMZRange) this.HiMZRange = SWATHRunner.HiMZ;
		if (SWATHRunner.WidthMZ > this.SWATHWidest) this.SWATHWidest = SWATHRunner.WidthMZ;
		if (SWATHRunner.WidthMZ < this.SWATHNarrowest) this.SWATHNarrowest = SWATHRunner.WidthMZ;
	        // thwop
		float TICSum = 0;
		while (SMRunner != null) {
		    SWATHRunner.MSMSCount++;
		    SWATHRunner.HiRT = SMRunner.ScanStartTime;
		    TICSum = TICSum + SMRunner.mzMLtic;
		    if (SMRunner.mzMLPeakCount > SWATHRunner.PkCountMax) SWATHRunner.PkCountMax = SMRunner.mzMLPeakCount;
		    SMRunner = SMRunner.Next;
		}
		SWATHRunner.TotalTIC = TICSum;
		if (SWATHRunner.TotalTIC < this.TotalTICMin) this.TotalTICMin = SWATHRunner.TotalTIC;
		if (SWATHRunner.TotalTIC > this.TotalTICMax) this.TotalTICMax = SWATHRunner.TotalTIC;
		SMRunner = TheseScans.Next;
		float TIC25 = TICSum / 4;
		float TIC50 = TICSum / 2;
		float TIC75 = TIC25+TIC50;
		float TICSoFar = 0;
		float TICAfterThisScan = 0;
		int[] PkCounts = new int[SWATHRunner.MSMSCount];
		int PkCountIndex = 0;
		float[] CycleTimes = new float[SWATHRunner.MSMSCount-1];
		float LastScanStartTime=0;
		if (SWATHRunner.MSMSCount < this.SWATHCycleCountMin) this.SWATHCycleCountMin = SWATHRunner.MSMSCount;
		if (SWATHRunner.MSMSCount > this.SWATHCycleCountMax) this.SWATHCycleCountMax = SWATHRunner.MSMSCount;
		while (SMRunner != null) {
		    TICAfterThisScan = TICSoFar + SMRunner.mzMLtic;
		    if ( (TIC25 > TICSoFar) && (TIC25 <= TICAfterThisScan) ) SWATHRunner.TIC25ileRT = SMRunner.ScanStartTime;
		    if ( (TIC50 > TICSoFar) && (TIC50 <= TICAfterThisScan) ) SWATHRunner.TIC50ileRT = SMRunner.ScanStartTime;
		    if ( (TIC75 > TICSoFar) && (TIC75 <= TICAfterThisScan) ) SWATHRunner.TIC75ileRT = SMRunner.ScanStartTime;
		    TICSoFar = TICAfterThisScan;
		    PkCounts[PkCountIndex] = SMRunner.mzMLPeakCount;
		    if (PkCountIndex > 0) CycleTimes[PkCountIndex-1] = SMRunner.ScanStartTime - LastScanStartTime;
		    PkCountIndex++;
		    LastScanStartTime = SMRunner.ScanStartTime;
		    SMRunner = SMRunner.Next;
		}
		if (SWATHRunner.TIC50ileRT < this.TIC50ileRTMin) this.TIC50ileRTMin = SWATHRunner.TIC50ileRT;
		if (SWATHRunner.TIC50ileRT > this.TIC50ileRTMax) this.TIC50ileRTMax = SWATHRunner.TIC50ileRT;
		Array.Sort(PkCounts);
		SWATHRunner.PkCount25ile = PkCounts[SWATHRunner.MSMSCount / 4];
		SWATHRunner.PkCount50ile = PkCounts[SWATHRunner.MSMSCount / 2];
		SWATHRunner.PkCount75ile = PkCounts[(SWATHRunner.MSMSCount / 4) + (SWATHRunner.MSMSCount/2)];
		SWATHRunner.PkCountMax = PkCounts[SWATHRunner.MSMSCount-1];
		if (SWATHRunner.PkCount50ile < this.PkCount50ileMin) this.PkCount50ileMin = SWATHRunner.PkCount50ile;
		if (SWATHRunner.PkCount50ile > this.PkCount50ileMax) this.PkCount50ileMax = SWATHRunner.PkCount50ile;
		Array.Sort(CycleTimes);
		//We multiply by 60 to get seconds from minutes.
		SWATHRunner.CycleTimeMedian = 60*CycleTimes[SWATHRunner.MSMSCount / 2];
		SumMedianCycleTime += SWATHRunner.CycleTimeMedian;
		TheseScans = ScansTable.ExciseAllAtFirstIsolationTarget();
	    }
	    this.MedianCycleTimeAverage = SumMedianCycleTime / (float)this.SWATHCount;
	}
	
	public void BubbleSortScans()
	{
	    /* Our goal is to get all the MS/MS scans of the same type
	     * sorted together, ordered by their retention times.
	     * Ideally, they are already sorted by retention time, but
	     * it's possible that they are not. */
	    ScanMetrics SMRunner;
	    ScanMetrics SMBuffer1;
	    ScanMetrics SMBuffer2;
	    bool MadeChanges = true;
	    while (MadeChanges)
	    {
		MadeChanges = false;
		SMRunner = this.ScansTable;
		while (SMRunner != null)
		{
		    SMBuffer1 = SMRunner.Next;
		    if (SMBuffer1 != null)
		    {
			SMBuffer2 = SMBuffer1.Next;
			if (SMBuffer2 != null)
			{
			    if (SMBuffer2.ComesBefore(SMBuffer1))
			    {
				Console.WriteLine("Swapping these two:");
				Console.WriteLine(SMBuffer1.Stringify());
				Console.WriteLine(SMBuffer2.Stringify());
				// Reorder the pair after SMRunner
				SMRunner.Next = SMBuffer2;
				SMBuffer1.Next = SMBuffer2.Next;
				SMBuffer2.Next = SMBuffer1;
				MadeChanges = true;
			    }
			}
		    }
		    SMRunner = SMRunner.Next;
		}
	    }
	}
	
        public void ReadFromMZML(XmlReader Xread)
        {
            /*
              Do not think of this code as a general-purpose mzML
              reader.  It is intended to populate only the fields that
              DIAMetric cares about.  For example, it entirely ignores
              the array of m/z values and intensities stored for any
              spectra.  This is intended to glean only the required
              fields in a single pass of the file.  It uses only the
              System.Xml libraries from Microsoft, obviating the need
              for any add-in libraries (to simplify the build process
              to something even I can use).
             */
            ScansRunner = ScansTable;
            while (Xread.Read())
            {
                var ThisNodeType = Xread.NodeType;
                if (ThisNodeType == XmlNodeType.Element)
                {
                    if (Xread.Name == "run")
                    {
                        /*
                          We directly read relatively little
                          information about the mzML file as a whole.
                          Here we grab the startTimeStamp, and
                          elsewhere we grab the file name root,
                          instrument model, and serial number.
                        */
                        StartTimeStamp = Xread.GetAttribute("startTimeStamp");
                    }
                    else if (Xread.Name == "spectrum")
                    {
                        /*
                          We only create a new ScanMetrics object if
                          it isn't an MS1 scan.  We need to keep two
                          pieces of information from this new spectrum
                          header in case we do make a new ScanMetrics
                          object.
                        */
                        var ThisPeakCount = Xread.GetAttribute("defaultArrayLength");
                        LastPeakCount = int.Parse(ThisPeakCount);
                        LastNativeID = Xread.GetAttribute("id");
                    }
                    else if (Xread.Name == "cvParam")
                    {
                        var Accession = Xread.GetAttribute("accession");
                        switch (Accession)
                        {
                            /*
                              If you see that instrument model is ever
                              blank for an mzML, there are two likely
                              causes.  The first would be that the
                              mzML converter has not listed the CV
                              term for the instrument type in the
                              mzML-- ProteoWizard _does_ record this
                              information.  The second likely cause is
                              that the CV term relating to your
                              instrument model is missing from this
                              list.  Just add a "case" line for it and
                              recompile.
                             */
                            case "MS:1000557":
			    case "MS:1000932":
			    case "MS:1001742":
                            case "MS:1001910":
                            case "MS:1001911":
                            case "MS:1002416":
                            case "MS:1002523":
			    case "MS:1002533":
			    case "MS:1002634":
                            case "MS:1002732":
			    case "MS:1002877":
			    case "MS:1003005":
			    case "MS:1003028":
                            case "MS:1003029":
			    case "MS:1003094":
			    case "MS:1003123":
			    case "MS:1003293":
                                Instrument = Xread.GetAttribute("name");
                                break;
                            case "MS:1000529":
                                SerialNumber = Xread.GetAttribute("value");
                                break;
                            case "MS:1000016":
                                var ThisStartTime = Xread.GetAttribute("value");
                                // We need the "InvariantCulture" nonsense because some parts of the world separate decimals with commas.
                                var ThisStartTimeFloat = Single.Parse(ThisStartTime, CultureInfo.InvariantCulture);
                                ScansRunner.ScanStartTime = ThisStartTimeFloat;
                                if (ThisStartTimeFloat > MaxScanStartTime) MaxScanStartTime = ThisStartTimeFloat;
                                break;
			    case "MS:1001581":
				var ThisFAIMS = Xread.GetAttribute("value");
				ScansRunner.FAIMScompVoltage = Single.Parse(ThisFAIMS, CultureInfo.InvariantCulture);
				break;
			    case "MS:1000827":
				var ThisIsolationTarget = Xread.GetAttribute("value");
				ScansRunner.IsolationTarget = Double.Parse(ThisIsolationTarget, CultureInfo.InvariantCulture);
				break;
			    case "MS:1000828":
				var ThisIsolationLower = Xread.GetAttribute("value");
				ScansRunner.IsolationLowerOffset = Double.Parse(ThisIsolationLower, CultureInfo.InvariantCulture);
				break;
			    case "MS:1000829":
				var ThisIsolationHigher = Xread.GetAttribute("value");
				ScansRunner.IsolationHigherOffset = Double.Parse(ThisIsolationHigher, CultureInfo.InvariantCulture);
				break;
			    case "MS:1000285":
				var ThisTIC = Xread.GetAttribute("value");
				ScansRunner.mzMLtic = Single.Parse(ThisTIC, CultureInfo.InvariantCulture);
				break;
                            case "MS:1000511":
                                var ThisLevel = Xread.GetAttribute("value");
                                var ThisLevelInt = int.Parse(ThisLevel);
                                if (ThisLevelInt == 1)
                                {
                                    // We do very little with MS scans other than count them.
                                    mzMLMS1Count++;
                                }
                                else
                                {
                                    /*
                                      If we detect an MS of level 2 or
                                      greater in the mzML, we have
                                      work to do.  Each MS/MS is
                                      matched by an item in the linked
                                      list of ScanMetrics for each
                                      LCMSMSExperiment.  We will need
                                      to capture some information we
                                      already saw (such as the
                                      NativeID of this scan) and set
                                      up for collection of some
                                      additional information in the
                                      Activation section.
                                     */
                                    mzMLMSnCount++;
                                    ScansRunner.Next = new ScanMetrics();
                                    ScansRunner = ScansRunner.Next;
                                    ScansRunner.NativeID = LastNativeID;
                                    ScansRunner.mzMLPeakCount = LastPeakCount;
                                    if (LastPeakCount > MaxPkCount)
                                        mzMLPeakCountDistn[MaxPkCount]++;
                                    else
                                        mzMLPeakCountDistn[LastPeakCount]++;
                                }
                                break;
                        }
                    }
                }
            }
        }
	
        public LCMSMSExperiment Find(string Basename)
        {
            /*
              We receive an msAlign filename.  We seek the
              LCMSMSExperiment in this linked list that has the
              corresponding filename.
            */
            var LRunner = this.Next;
            while (LRunner != null)
            {
                if (Basename == LRunner.SourceFile)
                    return LRunner;
                LRunner = LRunner.Next;
            }
            return null;
        }

        public void ParseScanNumbers()
        {
            /*
              Instrument manufacturers differ in the ways that they
              report the identities of each MS and MS/MS.  Because
              TopFD was created for Thermo instruments, though, it
              expects each spectrum to have a unique scan number for a
              given RAW file.  To match to TopFD msAlign files, we'll
              need to extract those from the NativeIDs.
             */
            var SRunner = this.ScansTable.Next;
            while (SRunner != null)
            {
                // Example of Thermo NativeID: controllerType=0 controllerNumber=1 scan=12 (ProteoWizard)
                // Example of SCIEX NativeID: sample=1 period=1 cycle=806 experiment=2 (ProteoWizard)
		// Example of SCIEX NativeID: sample=1 period=1 cycle=7207 experiment=1 (SCIEX MS Data Converter)
                // Example of Bruker NativeID: scan=55 (TIMSConvert)
                // Example of Bruker NativeID: merged=102 frame=13 scanStart=810 scanEnd=834 (ProteoWizard)
                var Tokens = SRunner.NativeID.Split(' ');
		foreach (var ThisTerm in Tokens)
		{
		    var Tokens2 = ThisTerm.Split('=');
		    if ( Tokens2[0].Equals("cycle") || Tokens2[0].Equals("scan") || Tokens2[0].Equals("scanStart") )
		    {
			SRunner.ScanNumber = int.Parse(Tokens2[1]);
		    }
		}
                SRunner = SRunner.Next;
            }
        }

        public ScanMetrics GoToScan(int Target)
        {
            var SRunner = this.ScansTable.Next;
            while (SRunner != null)
            {
                if (SRunner.ScanNumber == Target)
                    return SRunner;
                SRunner = SRunner.Next;
            }
            return null;
        }
	
        public int[] QuartilesOf(int[] Histogram)
        {
            var Quartiles = new int[5];
            var Sum = 0;
            int index;
            var AwaitingMin = true;
            for (index = 0; index < Histogram.Length; index++)
            {
                if (AwaitingMin && Histogram[index] > 0)
                {
                    AwaitingMin = false;
                    Quartiles[0] = index;
                }
                if (Histogram[index] > 0)
                    Quartiles[4] = index;
                Sum += Histogram[index];
            }
            var CountQ1 = Sum / 4;
            var CountQ2 = Sum / 2;
            var CountQ3 = CountQ1 + CountQ2;
            Sum = 0;
            for (index = 0; index < Histogram.Length; index++)
            {
                var ThisCount = Histogram[index];
                if (Sum < CountQ1 && CountQ1 <= Sum + ThisCount)
                    Quartiles[1] = index;
                if (Sum < CountQ2 && CountQ2 <= Sum + ThisCount)
                    Quartiles[2] = index;
                if (Sum < CountQ3 && CountQ3 <= Sum + ThisCount)
                    Quartiles[3] = index;
                Sum += ThisCount;
            }
            return Quartiles;
        }

        public void WriteTextQCReport()
        {
            /*
              We have two TSV outputs.  The "byRun" report contains a
              row for each LC-MS/MS (or mzML) in this directory.  The
              "byMSn" report contains a row for each MS/MS in each
              mzML in this directory.
             */
	    //TODO: Should I be reporting distribution of deconvolved precursor mass by RAW?
            var LCMSMSRunner = this.Next;
            const string delim = "\t";
            using (var TSVbyRun = new StreamWriter("DIAMetric-byRun.tsv"))
            {
                TSVbyRun.WriteLine("SourceFile\tInstrument\tSerialNumber\tStartTimeStamp\tRTDuration" +
				   "\tmzMLMS1Count\tmzMLMSnCount\tIsolationWindowCount\tCyclesMin\tCyclesMax" +
				   "\tMZRangeMin\tMZRangeMax\tIsolationWindowWidthMin\tIsolationWidowWidthMax" +
				   "\tAverageMedianCycleTime\tTICMedianRTMin\tTICMedianRTMax\tTotalTICMin\tTotalTICMax" +
				   "\tPkCountMedianMin\tPkCountMedianMax");
                while (LCMSMSRunner != null)
                {
                    // Actually write the metrics to the byRun file...
                    TSVbyRun.Write(LCMSMSRunner.SourceFile + delim);
                    TSVbyRun.Write(LCMSMSRunner.Instrument + delim);
                    TSVbyRun.Write(LCMSMSRunner.SerialNumber + delim);
                    TSVbyRun.Write(LCMSMSRunner.StartTimeStamp + delim);
                    TSVbyRun.Write(LCMSMSRunner.MaxScanStartTime + delim);
                    TSVbyRun.Write(LCMSMSRunner.mzMLMS1Count + delim);
                    TSVbyRun.Write(LCMSMSRunner.mzMLMSnCount + delim);
		    TSVbyRun.Write(LCMSMSRunner.SWATHCount + delim);
		    TSVbyRun.Write(LCMSMSRunner.SWATHCycleCountMin + delim);
		    TSVbyRun.Write(LCMSMSRunner.SWATHCycleCountMax + delim);
		    TSVbyRun.Write(LCMSMSRunner.LoMZRange + delim);
		    TSVbyRun.Write(LCMSMSRunner.HiMZRange + delim);
		    TSVbyRun.Write(LCMSMSRunner.SWATHNarrowest + delim);
		    TSVbyRun.Write(LCMSMSRunner.SWATHWidest + delim);
		    TSVbyRun.Write(LCMSMSRunner.MedianCycleTimeAverage + delim);
		    TSVbyRun.Write(LCMSMSRunner.TIC50ileRTMin + delim);
		    TSVbyRun.Write(LCMSMSRunner.TIC50ileRTMax + delim);
		    TSVbyRun.Write(LCMSMSRunner.TotalTICMin + delim);
		    TSVbyRun.Write(LCMSMSRunner.TotalTICMax + delim);
		    TSVbyRun.Write(LCMSMSRunner.PkCount50ileMin + delim);
		    TSVbyRun.WriteLine(LCMSMSRunner.PkCount50ileMax + delim);
		    /*
                    foreach (var ThisQuartile in LCMSMSRunner.mzMLPeakCountQuartiles)
                        TSVbyRun.Write(ThisQuartile + delim);
		    */
                    LCMSMSRunner = LCMSMSRunner.Next;
                }
            }
            LCMSMSRunner = this.Next;
            using (var TSVbySWATH = new StreamWriter("DIAMetric-byIsolationWindow.tsv"))
	    {
                TSVbySWATH.WriteLine("SourceFile\tLoMZ\tHiMZ\tWidthMZ\tFAIMS\tMSMSCount\tRTMin\tRTMax\tCycleTimeMedian\tTIC25ileRT\tTIC50ileRT\tTIC75ileRT\tTotalTIC\tPkCount25ile\tPkCount50ile\tPkCount75ile\tPkCountMax");
		while (LCMSMSRunner != null)
		{
		    var SWATHRunner = LCMSMSRunner.SWATHs.Next;
		    while (SWATHRunner != null)
		    {
			TSVbySWATH.Write(LCMSMSRunner.SourceFile + delim);
			TSVbySWATH.Write(SWATHRunner.LoMZ + delim);
			TSVbySWATH.Write(SWATHRunner.HiMZ + delim);
			TSVbySWATH.Write(SWATHRunner.WidthMZ + delim);
			TSVbySWATH.Write(SWATHRunner.FAIMS + delim);
			TSVbySWATH.Write(SWATHRunner.MSMSCount + delim);
			TSVbySWATH.Write(SWATHRunner.LoRT + delim);
			TSVbySWATH.Write(SWATHRunner.HiRT + delim);
			TSVbySWATH.Write(SWATHRunner.CycleTimeMedian + delim);
			TSVbySWATH.Write(SWATHRunner.TIC25ileRT + delim);
			TSVbySWATH.Write(SWATHRunner.TIC50ileRT + delim);
			TSVbySWATH.Write(SWATHRunner.TIC75ileRT + delim);
			TSVbySWATH.Write(SWATHRunner.TotalTIC + delim);
			TSVbySWATH.Write(SWATHRunner.PkCount25ile + delim);
			TSVbySWATH.Write(SWATHRunner.PkCount50ile + delim);
			TSVbySWATH.Write(SWATHRunner.PkCount75ile + delim);
			TSVbySWATH.WriteLine(SWATHRunner.PkCountMax);
			SWATHRunner = SWATHRunner.Next;
		    }
		    LCMSMSRunner = LCMSMSRunner.Next;
		}
	    }

        }
    }
}

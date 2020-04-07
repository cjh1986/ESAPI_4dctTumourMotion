using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using System.Collections;

namespace LungMotionPull
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                using (Application app = Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                Console.ReadLine();
            }
        }

        static void Execute(Application app)
        {
            //Intro
            Console.WriteLine("Standalone script to pull imaging data from 4DCT planning scan structure sets");

            //Open input text file for list of ID numbers.
            StreamReader myReader = new StreamReader(getFilePath(@"\\oncology-is\VA_DATA$\physicists\ESAPI Scripting\Patient lists"));
            System.Collections.ArrayList IDList = new System.Collections.ArrayList(); //empty array

            //Reads in ID number and assigns to array.
            string IDNo = myReader.ReadLine();
            while (IDNo != null)
            {
                IDList.Add(IDNo);
                IDNo = myReader.ReadLine();
            }

            //set output file location
            string outputDir = @"\\oncology-is\VA_DATA$\physicists\ESAPI Scripting\Reports";

            string filename = string.Format(@"{0}\LungMotionPhase_{1}.csv",
                outputDir, DateTime.Now.ToString("dd_MM_yyyy"));
            Console.WriteLine("Output location for Phase info:");
            Console.WriteLine(filename);

            string filename2 = string.Format(@"{0}\LungMotionPlan_{1}.csv",
            outputDir, DateTime.Now.ToString("dd_MM_yyyy"));
            Console.WriteLine("Output location for Plan info:");
            Console.WriteLine(filename2);

            //Creates file for phase information
            System.IO.StreamWriter myWriterPhase = new StreamWriter(filename); //creates empty .csv
            myWriterPhase.AutoFlush = true;

            //Writes header for csv for phases
            myWriterPhase.Write("ID, Course, Plan, Protocol, Dose, Plan creation, Structure Set, 4D?, Phase, GTVvolume (cc)," +
            "x-lat (cm),y-vrt (cm),z-lng (cm), " +
            "x-lat2Lung (cm),y-vrt2Lung (cm),z-lng2Lung (cm), " +
            "x-lat2PTV (cm),y-vrt2PTV (cm),z-lng2PTV (cm), " +
            "PTV Vol(cc), lungVol(cc), key");
            myWriterPhase.Write("\n");

            //Creates file for summary information for each plan.
            System.IO.StreamWriter myWriterPlan = new StreamWriter(filename2); //creates empty .csv
            myWriterPhase.AutoFlush = true;

            //Writes header for csv for plan.
            myWriterPlan.Write("Key, ID, Course, Site, Plan, Protocol, Dose, Plan creation, Structure Set, " +
            "PTV Vol (cc), lungVol (cc), 4D?,  GTV mean vol (cc), " +
            "x-lat range (cm),y-vrt range (cm), z-lng range (cm), " +
            "x-lat2Lung (cm), y-vrt2Lung (cm), z-lng2Lung (cm), " +
            "x-lat2PTV (cm), y-vrt2PTV (cm), z-lng2PTV (cm), Laterality,");
            myWriterPlan.Write("\n");


            double Count = 0; //counts patients.
            foreach (string id in IDList) //runs through list of IDs
            {
                Count++;
                Console.WriteLine("Processing patient {0} of {1}, id: {2}", Count, IDList.Count, id);

                Patient pat = app.OpenPatientById(id);

                foreach (var course in pat.Courses) //Runs through all courses
                {
                    //filters for relevant course
                    if (course.Id.ToUpper().Contains("QA")
                        || course.Id.ToUpper().Contains("QC")
                        || course.Id.ToUpper().Contains("XIO")
                        || !course.Id.ToUpper().Contains("LUNG"))
                    {
                        //skips irrelevant courses for loop
                        continue;
                    }

                    foreach (var ps in course.PlanSetups) //Runs through all plans
                    {
                        // Checks for arcs.
                        int arcs = 0;
                        for (int Beams = 0; Beams < ps.Beams.Count(); Beams++) 
                        {
                            if ((ps.Beams.ElementAt(Beams).IsSetupField)) { continue; }
                            if (ps.Beams.ElementAt(Beams).Technique.Id.ToUpper().Contains("ARC"))
                            { arcs++; } //adds to index if criteria are met.          
                        }

                        string Approval = ps.ApprovalStatus.ToString();
                        //Filters plans
                        if (!(ps.IsDoseValid)                                 // skip plans with no dose
                            // || Approval.Contains("Retired")                // skips retired
                            || Approval.Contains("Rejected")                  // skips rejected
                            // || !(Approval.Contains("TreatmentApproved"))   // only includes treatment approved plans
                            || ps.Id.ToUpper().Contains("QA")                 // skip QA plans
                            || ps.Id.ToUpper().Contains("SCRIPT")             // skip plans with script in ID
                            || ps.Id.ToUpper().Contains("#")
                            //|| !ps.Id.ToUpper().Contains("SABR")
                            || ps.StructureSet == null                        // no structure set
                            || arcs == 0                                      // no arcs
                            )
                        { continue; }

                        if (ps.IsTreated)
                        {
                            //placed in a try catch loop to catch errors but not end code.
                            try
                            {
                                //Passes data to method to work out results.
                                DataStruc tempData = processPlan(pat, ps);
                                //Writes phases information to csv
                                csvWritePhase(tempData, myWriterPhase, filename);
                                //Writes plan information to csv
                                csvWritePlan(tempData, myWriterPlan, filename2);
                            }
                            catch (Exception e)
                            {
                                //outputs error to console.
                                Console.Error.WriteLine(e.ToString());
                            }
                        }
                    }
                }
                app.ClosePatient();
            }

            myWriterPhase.Close();
            myWriterPlan.Close();

            Console.WriteLine("All patients processed. Hit Return to open results and exit console");
            Console.ReadLine();
            System.Diagnostics.Process.Start(filename);   //opens results spreadsheet      
            System.Diagnostics.Process.Start(filename2);   //opens results spreadsheet 
        }

        public static DataStruc processPlan(Patient pat, PlanSetup ps)
        {
            //Sets the image FOR which is unique to images acquired at the same time.
            string FOR = ps.StructureSet.Image.FOR;
            StructureSet mainSS = ps.StructureSet;

            //Makes place holder for results.
            DataStruc Result = new DataStruc();

            //Produces generic plan results.
            Result.ID = pat.Id;
            Result.SSiD = mainSS.Id;
            Result.PD = ps.TotalDose.Dose;
            Result.protocol = ps.ProtocolID;
            Result.plan = ps.Id;
            Result.plannedOn = ps.HistoryDateTime;
            Result.Site = SiteFind(ps.ProtocolID, ps.Course.Id);
            Result.course = ps.Course.Id;
            Result.planKey = ps.UID + ps.Id;
            Result.fourD = "No";

            //Pre-creates list to assign phase data to.
            Result.GTVPhVol = new List<double>();
            Result.phase = new List<string>();
            Result.xGTVph = new List<double>();
            Result.yGTVph = new List<double>();
            Result.zGTVph = new List<double>();
            Result.xGTVph2Lung = new List<double>();
            Result.yGTVph2Lung = new List<double>();
            Result.zGTVph2Lung = new List<double>();
            Result.xGTVph2PTV = new List<double>();
            Result.yGTVph2PTV = new List<double>();
            Result.zGTVph2PTV = new List<double>();
            Result.key = new List<string>();

            //Creates relevant structures using linq.
            Structure Lungs = null;
            Lungs = mainSS.Structures.Where(x => x.Id.ToUpper() == "LUNGS").FirstOrDefault();
            if (Lungs == null)
            { Lungs = mainSS.Structures.Where(x => x.Id.ToUpper() == "WHOLE LUNG").FirstOrDefault(); }

            if (Lungs == null)
            { Lungs = mainSS.Structures.Where(x => x.Id.ToUpper() == "LUNGS-GTV").FirstOrDefault(); }

            if (Lungs != null)
            { Result.LungVol = Lungs.Volume; }
            else
            {
                Result.LungVol = 0;
            }

            Structure PTV = null;
            PTV = mainSS.Structures.Where(x => x.Id.ToUpper() == "PTV").FirstOrDefault();

            if (PTV == null)
            { PTV = mainSS.Structures.Where(x => x.Id.ToUpper().Contains("PTV")).FirstOrDefault(); }

            Result.PTVVol = PTV.Volume;

            //Makes a collection of 4DCT phases, not including CBCT.
            var SSColl = pat.StructureSets.Where(x => x.Image.FOR == FOR
                && !x.Id.ToUpper().Contains("CBCT")
                && !x.Id.ToUpper().Contains("INTERPLAY")
                && x.Id.ToUpper().Contains("CT")); 

            if (SSColl != null)
            {
                //Cycle through phases to find CoM and volumes.
                foreach (var SS in SSColl)
                {
                    //Cycles through structures in structureset.
                    foreach (var S in SS.Structures)
                    {
                        if (S.Id.ToUpper().Equals("GTV PH") && S.IsEmpty != true) // finds the centre of the GTV
                        {
                            //Commit phase ranges to list.
                            Result.xGTVph.Add(S.CenterPoint.x);
                            Result.yGTVph.Add(S.CenterPoint.y);
                            Result.zGTVph.Add(S.CenterPoint.z);
                            Result.GTVPhVol.Add(S.Volume);
                            Result.phase.Add(SS.Id);
                            //Adds to primary key if phases are present.
                            Result.key.Add(ps.UID + SS.Id);

                            //Determines offset between PTV and phase. +1000 removes the impact of sign changes.
                            Result.xGTVph2PTV.Add((1000 + PTV.CenterPoint.x) - (1000 + S.CenterPoint.x));
                            Result.yGTVph2PTV.Add((1000 + PTV.CenterPoint.y) - (1000 + S.CenterPoint.y));
                            Result.zGTVph2PTV.Add((1000 + PTV.CenterPoint.z) - (1000 + S.CenterPoint.z));

                            if (Lungs != null)
                            {
                                Result.xGTVph2Lung.Add((1000 + Lungs.CenterPoint.x) - (1000 + S.CenterPoint.x));
                                Result.yGTVph2Lung.Add((1000 + Lungs.CenterPoint.y) - (1000 + S.CenterPoint.y));
                                Result.zGTVph2Lung.Add((1000 + Lungs.CenterPoint.z) - (1000 + S.CenterPoint.z));
                            }
                            else
                            {
                                Result.xGTVph2Lung.Add(0);
                                Result.yGTVph2Lung.Add(0);
                                Result.zGTVph2Lung.Add(0);
                            }

                        }
                    }

                }                    
                //Reduces the GTV phase co-ordinate to be relative to zero.
                    if (Result.xGTVph.Count >= 1)
                    {

                    Result.fourD = "Yes";
                        double minX = Result.xGTVph.Min();
                        double minY = Result.yGTVph.Min();
                        double minZ = Result.zGTVph.Min();

                        for (int i = 0; i < Result.key.Count(); i++)
                        {
                            Result.xGTVph[i] = (1000 + Result.xGTVph[i]) - (1000 + minX); //1000 + removes problem of going over the axis, result is relative anyway.
                            Result.yGTVph[i] = (1000 + Result.yGTVph[i]) - (1000 + minY);
                            Result.zGTVph[i] = (1000 + Result.zGTVph[i]) - (1000 + minZ);
                        } 
                    }
            }
            else
            {
                Result.key.Add(ps.UID + ps.Id);
            }

            return Result; //returns result.
        }

        public static string getFilePath()
        {
            //helper method which opens a file browser to allow user to select input
            //default starting pathway is IOFolder as below

            string filePathway = "";

            System.Windows.Forms.OpenFileDialog open1 = new System.Windows.Forms.OpenFileDialog();
            open1.InitialDirectory = @"\\oncology-is\va_data$\physicists\ESAPI Scripting\Patient lists\";
            open1.Title = "Please open the input file";

            if (open1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePathway = open1.FileName;
            }
            return filePathway;
        }

        public static string getFilePath(string initialPath)
        {
            //overloaded method. Initial browsing directory set by string initialPath
            string filePathway = "";
            System.Windows.Forms.OpenFileDialog open1 = new System.Windows.Forms.OpenFileDialog();
            open1.InitialDirectory = initialPath;
            open1.Title = "Please open the input file";
            if (open1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePathway = open1.FileName;
            }
            return filePathway;
        }

        public static void csvWritePhase(DataStruc Result, StreamWriter myWriterPhase, string filename)
        {
            for (int i = 0; i < Result.key.Count(); i++)
            {
                myWriterPhase.Write(
                Result.ID + "," +
                Result.course + "," +
                Result.plan + "," +
                Result.protocol + "," +
                Result.PD + "," +
                Result.plannedOn + "," +
                Result.SSiD + "," +
                Result.fourD + "," +
                Result.phase[i] + "," +
                Result.GTVPhVol[i] + "," +
                Result.xGTVph[i] / 10 + "," +
                Result.yGTVph[i] / 10 + "," +
                Result.zGTVph[i] / 10 + "," +
                Result.xGTVph2Lung[i] / 10 + "," +
                Result.yGTVph2Lung[i] / 10 + "," +
                Result.zGTVph2Lung[i] / 10 + "," +
                Result.xGTVph2PTV[i] / 10 + "," +
                Result.yGTVph2PTV[i] / 10 + "," +
                Result.zGTVph2PTV[i] / 10 + "," +
                Result.PTVVol + "," +
                Result.LungVol + "," +
                Result.key[i]);
                myWriterPhase.Write("\n");
            }
        }

        public static void csvWritePlan(DataStruc Result, StreamWriter myWriterPlan, string filename)
        {
            //Summarises the 4D information per plan into 1 line.

            string output =
                Result.planKey + "," +
                Result.ID + "," +
                Result.course + "," +
                Result.Site + "," +
                Result.plan + "," +
                Result.protocol + "," +
                Result.PD + "," +
                Result.plannedOn + "," +
                Result.SSiD + "," +
                Result.PTVVol + "," +
                Result.LungVol + "," +
                Result.fourD;

            //Only includes 4D data if it was found.
            if (Result.GTVPhVol.Count() != 0)
            {
                //Determines laterality from tumour co-ordinates.
                string Laterality = "Unknown";
                if (Result.xGTVph2Lung.Average() < -20 && Result.zGTVph2Lung.Average() < 0)
                {
                    Laterality = "Left Upper";
                }
                else if (Result.xGTVph2Lung.Average() < -20 && Result.zGTVph2Lung.Average() >= 0)
                {
                    Laterality = "Left Lower";
                }
                else if (Result.xGTVph2Lung.Average() > 20 && Result.zGTVph2Lung.Average() < 0)
                {
                    Laterality = "Right Upper";
                }
                else if (Result.xGTVph2Lung.Average() > 20 && Result.zGTVph2Lung.Average() >= 0)
                {
                    Laterality = "Right Lower";
                }
                else
                {
                    if (Result.zGTVph2Lung.Average() < 0)
                    {
                        Laterality = "Central Upper";
                    }
                    else if (Result.zGTVph2Lung.Average() > 0)
                    {
                        Laterality = "Central Lower";
                    } 
                }


                output = output + "," +
                Result.GTVPhVol.Average() + "," +
                Result.xGTVph.Max() / 10 + "," +
                Result.yGTVph.Max() / 10 + "," +
                Result.zGTVph.Max() / 10 + "," +
                Result.xGTVph2Lung.Average() / 10 + "," +
                Result.yGTVph2Lung.Average() / 10 + "," +
                Result.zGTVph2Lung.Average() / 10 + "," +
                Result.xGTVph2PTV.Average() / 10 + "," +
                Result.yGTVph2PTV.Average() / 10 + "," +
                Result.zGTVph2PTV.Average() / 10 + "," +
                Laterality;
            }

            //Outputs string.
            myWriterPlan.Write(output);
            myWriterPlan.Write("\n");

        }

        private static string SiteFind(string Protocol, string Course)
        {
            //This block converts specificed protocol names to a generic site name for simplicity. 
            string Return = "";
            if (Protocol.ToUpper().Contains("H&N") || Protocol.ToUpper().Contains("HEAD&NECK") || Protocol.ToUpper().Contains("COMPARE"))
            {
                Return = "H&N";
            }
            else if (Protocol.ToUpper().Contains("LUNG") && !Protocol.ToUpper().Contains("SABR"))
            {
                Return = "Lung";
            }
            else if (Protocol.ToUpper().Contains("SABR"))
            {
                Return = "SABR";
            }
            else if (Protocol.ToUpper().Contains("OESO"))
            {
                Return = "Oesophagus";
            }
            else if (Protocol.ToUpper().Contains("PROSTATE") && !Protocol.ToUpper().Contains("NODES"))
            {
                Return = "Prostate";
            }
            else if (Protocol.ToUpper().Contains("PROSTATE&NODES"))
            {
                Return = "Prostate&Nodes";
            }
            else if (Protocol.ToUpper().Contains("RECTUM") || Protocol.ToUpper().Contains("ANAL"))
            {
                Return = "Rectum_anal";
            }
            else if (Protocol.ToUpper().Contains("UTERUS") || Protocol.ToUpper().Contains("CERVIX") || Protocol.ToUpper().Contains("VAGINA"))
            {
                Return = "Gynae";
            }
            else if (Protocol.ToUpper().Contains("BRAIN"))
            {
                Return = "Brain";
            }
            else
            //If a correct match can't be found, code just returns the protocol name.
            {
                if (Course.ToUpper().Contains("H&N") || Course.ToUpper().Contains("HEAD&NECK") || Course.ToUpper().Contains("COMPARE")
                    || Course.ToUpper().Contains("TONSIL")
                    || Course.ToUpper().Contains("NECK"))

                {
                    Return = "H&N";
                }
                else if (Course.ToUpper().Contains("LUNG") && !Course.ToUpper().Contains("SABR"))
                {
                    Return = "Lung";
                }
                else if (Course.ToUpper().Contains("SABR"))
                {
                    Return = "SABR";
                }
                else if (Course.ToUpper().Contains("PANC"))
                {
                    Return = "Pancrease";
                }
                else if (Course.ToUpper().Contains("STOM"))
                {
                    Return = "Stomach";
                }
                else if (Course.ToUpper().Contains("OESO"))
                {
                    Return = "Oesophagus";
                }
                else if (Course.ToUpper().Contains("SPINE"))
                {
                    Return = "Spine";
                }
                else if (Course.ToUpper().Contains("PROSTATE") && !Course.ToUpper().Contains("NODES"))
                {
                    Return = "Prostate";
                }
                else if (Course.ToUpper().Contains("PROSTATE&NODES"))
                {
                    Return = "Prostate&Nodes";
                }
                else if (Course.ToUpper().Contains("RECTUM") || Course.ToUpper().Contains("ANAL"))
                {
                    Return = "Rectum_anal";
                }
                else if (Course.ToUpper().Contains("UTERUS") || Course.ToUpper().Contains("CERVIX") || Course.ToUpper().Contains("VAGINA"))
                {
                    Return = "Gynae";
                }
                else if (Course.ToUpper().Contains("BRAIN"))
                {
                    Return = "Brain";
                }
                else
                {
                    {
                        Return = "Unknown";
                    }
                }
            }
            return Return;
        }


        public struct DataStruc
        {
            public string planKey;
            public string ID;
            public string course;
            public string plan;
            public double PD;
            public DateTime plannedOn;
            public string protocol;
            public List<string> key;                  //Unique Key used for database. UID+regtype
            public double LungVol;                    //lung vol
            public double PTVVol;
            public string SSiD;                       //Structure set ID
            public List<double> GTVPhVol;
            public string Site;
            public List<string> phase;
            public List<double> xGTVph;
            public List<double> yGTVph;
            public List<double> zGTVph;
            public List<double> xGTVph2PTV;
            public List<double> yGTVph2PTV;
            public List<double> zGTVph2PTV;
            public List<double> xGTVph2Lung;
            public List<double> yGTVph2Lung;
            public List<double> zGTVph2Lung;
            public string fourD;

        }
    }
}



///////////////////////////////////////////////////
// DVHDoseMetrics.cs
//
// Looking up Calculated DVH Metrics
//
// Applies to Eclipse V13.7
// Built by Matthew Schmidt 
// for questions: (matthew.schmidt@wustl.edu)
// Current calculations
//  - GEUD - setup up as a dictionary of structure IDs and a parameters
//  - Homogeneity Index - For any structure with PTV in the ID
//
//
//DVHDoseMetrics Copyright(c) 2019 Washington University. 
//Matthew Schmidt (matthew.schmidt@wustl.edu). Washington University hereby grants to you a non-transferable, non-exclusive, 
//royalty-free, non-commercial, research license to use and copy the computer code that may be downloaded 
//within this site (the “Software”).  You agree to include this license and the above copyright notice in all copies of the Software.  
//The Software may not be distributed, shared, or transferred to any third party.  
//This license does not grant any rights or licenses to any other patents, copyrights, 
//or other forms of intellectual property owned or controlled by Washington University.  
//If interested in obtaining a commercial license, please contact Washington University's Office of Technology Management (otm@wustl.edu).

 

//YOU AGREE THAT THE SOFTWARE PROVIDED HEREUNDER IS EXPERIMENTAL AND IS PROVIDED “AS IS”, 
//WITHOUT ANY WARRANTY OF ANY KIND, EXPRESSED OR IMPLIED, INCLUDING WITHOUT LIMITATION WARRANTIES 
//OF MERCHANTABILITY OR FITNESS FOR ANY PARTICULAR PURPOSE, OR NON-INFRINGEMENT OF ANY THIRD-PARTY PATENT, 
//COPYRIGHT, OR ANY OTHER THIRD-PARTY RIGHT.  IN NO EVENT SHALL THE CREATORS OF THE SOFTWARE OR WASHINGTON UNIVERSITY BE LIABLE 
//FOR ANY DIRECT, INDIRECT, SPECIAL, OR CONSEQUENTIAL DAMAGES ARISING OUT OF OR IN ANY WAY CONNECTED WITH THE SOFTWARE, 
//THE USE OF THE SOFTWARE, OR THIS AGREEMENT, WHETHER IN BREACH OF CONTRACT, TORT OR OTHERWISE, 
//EVEN IF SUCH PARTY IS ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.

 
//***************
//Note from the developer:
// Feel free to test out the functionality of this plug-in
//The gEUD is calculated as the sum(vi*Di^a)^(1/a)
//The HI is calculated as (D2-D98)/Dprescribed.
//This should work for both plans and plansums
//  Only one plansum should be open at a time
//  The script assumes that a plansum is created with the same structure set as it only takes the first structureset
//
//There may be some descrepencies in the gEUD calculation with Eclipse.
//This may be normal as taking high dose values to high power exponents leads to very large numbers.
//When evaluating this please look at the percentage different in Eclipse as the absolute difference may be misleading for high dose prescriptions.
//Please raise an issue on github if doses are different by ~>1-2%

using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace VMS.TPS
{
    public class Script
    {
        //TODO Change this Dictionary to Structure ID strings and a-values you would like to use.
        //the string search does a ToUpper on both sides so no need to worry about case.
        //beware of duplicate, i.e. Lung and Ipsilateral Lung could pick up the same structure twice.
        //The search strings here are on a Contains() method.
        Dictionary<string, double> structure_ids = new Dictionary<string, double>()
        {
            { "Heart", 0.5 },
            { "Cord",20 },
            { "Parotid",0.5 },
            { "Lung",0.5 },
            { "Bladder",0.5 },
            { "Rectum", 20 },
            {"stem",20 },
            {"PTV",-0.1 }
        };

        public Script()
        {
        }
        PlanningItem pi;
        StructureSet ss;
        List<DoseInfo> di_list = new List<DoseInfo>();
        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {
            // TODO : Add here your code that is called when the script is launched from Eclipse
            PlanSetup ps = context.PlanSetup;
            var psums = context.PlanSumsInScope;
            if (ps == null && psums == null)//check to make sure a plan or plansum is open.
            {
                MessageBox.Show("Please open a plan or plansum before running this script");
                return;
            }
            if (psums.Count() > 1)//check to see if only 1 plansum is open. If more than 1, issue a warning, but move on.
            {
                MessageBox.Show("This script will work best with only one plansum open at a time. Please open only 1 plansum");
            }
            //default to the planning item being a plansetup, but if it doesn't exist make it the first plansum.
            pi = ps == null ? (psums.First() as PlanningItem) : (ps as PlanningItem);
            //if this is a plansum, assume the first structureset is the actual structure set.
            ss = ps == null ? psums.First().PlanSetups.First().StructureSet : ps.StructureSet;
            //loop through structures.
            foreach (Structure s in ss.Structures.Where(x => x.DicomType != "MARKER"))//filter out marker types
            {
                if (!s.IsEmpty)//check for empty structures
                {
                    foreach (string id in structure_ids.Keys)//search through dictionary
                    {
                        if (s.Id.ToUpper().Contains(id.ToUpper()))
                        {
                            //custom class is defined at the bottom.
                            di_list.Add(new DoseInfo
                            {
                                StructureId = s.Id,
                                a_parameter = structure_ids[id],
                                gEUD = CalculateGEUD(s, pi, structure_ids.FirstOrDefault(x => x.Key == id)),
                                HI = CalculateHI(s, pi),
                            });
                        }
                    }
                }
            }
            //write out dose info information to a messagebox.
            MessageBox.Show(String.Join("\n", di_list.Select(x => new
            {
                s = String.Format("Structure: {0}; gEUD = {1:F2} for a = {2}; HI = {3:F2}",
                 x.StructureId, x.gEUD, x.a_parameter, x.HI)
            })
                .Select(x => x.s)));
            //**Notes for developers.
            //the line above can be written as below if using a binary plug-in
            //MessageBox.Show(String.Join("\n", di_list.Select(x => new
            //{
            //    s = $"Structure: {x.StructureId}; GEUD: {x.gEUD:F2}; for a = {x.a_parameter}{(x.hi != 0 ? $" hi = {x.hi}" : "")}"
            //}).Select(x => x.s)));
            //if unfamiliar with linq, the line above can be written in the following way.
            //string output_s = "";
            //foreach(DoseInfo di in di_list)
            //{
            //    output_s += String.Format("Structure: {0}; gEUD = {1:F2} for a = {2}; HI = {3:F2}\n",
            //     di.StructureId, di.gEUD, di.a_parameter, di.HI);
            //}
            //MessageBox.Show(output_s);
        }
        //Homogeneity Index: An objective tool for assessment of conformal radiation treatment.
        //J Med Phys 2012 Oct-Dec; 37(4): 207-213
        //HI = (D2-D98)/Dpx100 where "D2 = minimum dose to 2% of the target volume...
        //D98 = minimum dose to 98% of the target volume...
        //and Dp = prescribed dose."
        private double CalculateHI(Structure s, PlanningItem pi)
        {
            //hi only needs to be calculated for ptv, so filter those out here.
            if (!s.Id.ToUpper().Contains("PTV"))
            {
                return Double.NaN;
            }
            //now check if the planning item is a plansetup or sum.
            if (pi is PlanSetup)
            {
                //plansetups have a method called GetDoseAtVolume.
                if (pi.Dose == null)
                {
                    MessageBox.Show("Plan has no dose");
                    return Double.NaN;
                }
                double d2 = (pi as PlanSetup).GetDoseAtVolume(s, 2, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
                double d98 = (pi as PlanSetup).GetDoseAtVolume(s, 98, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
                double hi = ((d2 - d98) / (pi as PlanSetup).TotalPrescribedDose.Dose) * 100;
                return hi;
            }
            else if (pi is PlanSum)
            {
                //must manually calculate value from DVH
                DVHData dvh = pi.GetDVHCumulativeData(s, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
                if (dvh == null)
                {
                    MessageBox.Show("Could not collect DVH");
                    return Double.NaN;
                }
                double d98 = dvh.CurveData.FirstOrDefault(x => x.Volume <= 98).DoseValue.Dose;
                double d2 = dvh.CurveData.FirstOrDefault(x => x.Volume <= 2).DoseValue.Dose;
                List<double> rx_doses = new List<double>();
                foreach (PlanSetup ps in (pi as PlanSum).PlanSetups)
                {
                    try
                    {
                        rx_doses.Add(ps.TotalPrescribedDose.Dose);
                    }
                    catch
                    {
                        MessageBox.Show("One of the prescriptions for the plansum is not defined");
                        return Double.NaN;
                    }
                }
                double rx = rx_doses.Sum();
                double hi = ((d2 - d98) / rx) * 100;
                return hi;
            }
            else
            {
                MessageBox.Show("Plan not handled correctly");
                return Double.NaN;
            }
        }
        //AAPM Report NO. 166: 
        //The use and QA of Biologically Related Models for Treatment Planning
        //gEUD = (sum(viDi^a))^1/a
        //"where vi is the fractional organ volume receiving a dose Di
        //and a is a tissue-specific parameter that describes the volume affect."
        private double CalculateGEUD(Structure s, PlanningItem pi, KeyValuePair<string, double> a_lookup)
        {
            //collect the DVH
            //if volume is not relative, make sure to normalize over the total volume during geud calculation.
            //double volume = s.Volume;
            //remember plansums must be absolute dose.
            DVHData dvh = pi.GetDVHCumulativeData(s, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            if (dvh == null)
            {
                MessageBox.Show("Could not calculate DVH");
                return Double.NaN;
            }
            //we need to get the differential volume from the definition. Loop through Volumes and take the difference with the previous dvhpoint
            double running_sum = 0;
            int counter = 0;
            foreach (DVHPoint dvhp in dvh.CurveData.Skip(1))
            {
                //volume units are in % (divide by 100)
                double vol_diff = Math.Abs(dvhp.Volume - dvh.CurveData[counter].Volume) / 100;
                double dose = dvhp.DoseValue.Dose;
                running_sum += vol_diff * Math.Pow(dose, a_lookup.Value);
                counter++;
            }
            double geud = Math.Pow(running_sum, 1 / a_lookup.Value);
            return geud;
        }

        public class DoseInfo
        {
            public string StructureId { get; set; }
            public double gEUD { get; set; }
            public double a_parameter { get; set; }
            public double HI { get; set; }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OCP_2._0
{
    public class TypesofAgreement : ObservableCollection<string>
    {
        List<object> titleIDs = new List<object>();

        public TypesofAgreement()
        {
            String titleSQL = "SELECT [TYPE_OF_AGREEMENT] FROM [REVINT]." + (new Contract()).getSchema() + ".[OCP_Contracts] WHERE NOT [TYPE_OF_AGREEMENT] IS NULL GROUP BY [TYPE_OF_AGREEMENT]";
                //"SELECT TYPE_OF_AGREEMENT FROM [REVINT].[dbo].[OCP_TypesOfAgreement]";

            new idMaker(titleSQL, titleIDs);

            foreach(object status in titleIDs)
            {
                Add(status.ToString());
            }
        }
    }
    public class Statuses : ObservableCollection<string>
    {
        List<object> titleIDs = new List<object>();

        public Statuses()
        {
                //"SELECT STATUS FROM [REVINT].[dbo].[OCP_Statuses]";
            new idMaker("SELECT [STATUS] FROM [REVINT]." + (new Contract()).getSchema() + ".[OCP_Contracts] WHERE NOT [STATUS] IS NULL GROUP BY [STATUS]", titleIDs);

            var itemsToRemove = this.ToList();

            foreach (var itemToRemove in itemsToRemove)
            {
                Remove(itemsToRemove.ToString());
            }

            foreach (object status in titleIDs)
            {
                Add(status.ToString());
            }
        }
        public void resetStatuses(String statusSQL)
        {
            titleIDs = new List<object>();
            new idMaker("SELECT [STATUS]" + statusSQL + "AND NOT [STATUS] IS NULL GROUP BY [STATUS]", titleIDs);

            var itemsToRemove = this.ToList();

            //this

            for (int i = this.Count - 1; i >= 0; i--)
            {
                RemoveAt(i);
            }

            foreach (object status in titleIDs)
            {
                Add(status.ToString());
            }
        }
    }
    public class UIDepartments : ObservableCollection<string>
    {
        List<object> titleIDs = new List<object>();

        public UIDepartments()
        {
            String titleSQL = "SELECT [UI_DEPARTMENT] FROM [REVINT]." + (new Contract()).getSchema() + ".[OCP_Contracts] WHERE NOT [UI_DEPARTMENT] IS NULL GROUP BY [UI_DEPARTMENT]";
                //"SELECT UI_DEPARTMENT FROM [REVINT].[dbo].[OCP_UIDepartments]";

            new idMaker(titleSQL, titleIDs);

            foreach (object status in titleIDs)
            {
                Add(status.ToString());
            }
        }
    }
    public class ReimbursementPeriods : ObservableCollection<string>
    {
        public ReimbursementPeriods()
        {
            Add("Per Hour");
            Add("Per Day");
            Add("Per Month");
            Add("Per Year");
            Add("Per Session");
            Add("Per Procedure");
            Add("Other");
        }
    }

    public class Clients : ObservableCollection <string>
    {
        private String clients = "Advanced Diagnostic Imaging;Anesthesia Resource Network;Athletic Department;Cedar Rapids Clinic for Women;Cedar Rapids Medical Education Foundation;Central Iowa Hospital Corporation d/b/a Blank Children's Hospital;Children's Heart Center;Clear Creek Amana Community Schools;College Community School District;Community Mental Health Center;Department of Corrections;Des Moines Iowa Methodist Medical Center;Dr. Chirantan Ghosh;Dr. Dewi Abramhoff;Dr. George S. Noble‚ Pediatric Surgeon affiliated with Des Moines Mercy;Dr. James Hopkins‚ Pediatric Surgeon affiliated with Des Moines Mercy;Dr. Kyriacos Panayides‚ Pediatric Surgeon;Dr. Thomas Fagan (Denver Children's Hospital);Eastern Iowa Health Center;Emma Goldman Clinic;Floyd County Medical Center;Fox Eye Surgery‚ LLC;Genesis Health System d/b/a Genesis Medical Center-Davenport;Genesis Health System d/b/a Genesis Medical Center-Illini;Genesis Heatlh System;Great River Health System;Great River Medical Center;Iowa City Cancer Treatment Center;Iowa City Community School District;Iowa City Hospice;Iowa Department of Public Health‚ State Medical Examiners Office;Keokuk Area Hospital;Keokuk County Health Center;Keokuk Health System;Lake Monrose Anesthesia Associates;Magellan;Mary Greeley Medical Center;Mason City Clinic;MECCA;Medical Associates‚ P.C.;Mercy - Springfield;Mercy Iowa City;Mercy Medical Center - Cedar Rapids;Mercy Medical Center - Centerville;Mercy Medical Center - Clinton;Mercy Medical Center - North Iowa;Mercy Medical Center Des Moines;Mercy Physician Associates;Mideast Iowa Mental Health/Community Mental Health Center;Mount Clemens Regional Medical Center;Naval Medical Center;Neurology Consultants‚ PC;New Client Name;North Dakota Department of Health;Northwest ENT;Oncology Associates;Pediatric Cardiology‚ P.C.;Pediatric Group Associates‚ S.C.;Perinatal Center of Iowa and Mercy Medical Center-Des Moines;Radiation Therapy Center of the Quad Cities;Regina Schools;ResCare‚ Inc;Sanford;Seasons Center - Spencer‚ IA;Skiff Hospital;Southeastern Iowa Medical Services‚ Inc;St. Lukes Hospital;State of Louisiana Department of Health and Hospitals;Sumner Community Club d/b/a Community Memorial Hospital;The Institute of Innovative Technology in Medical Education;The Queen's Medical Center;The Shelter House aka Fairweather Lodge;Trinity Regional Health System;University of Iowa Community Medical Services‚ Inc;University of Wisconsin Hospitals and Clinics Authority;University of Wisconsin Medical Foundation‚ Inc.;Van Buren County Hospital;Virginia Gay Hospital;Washington County Hospital;Wheaton Franciscan Healthcare;Wright Medical Center";
        List<object> titleIDs = new List<object>();
        String titleSQL = "SELECT [CLIENT] FROM [REVINT]."+(new Contract()).getSchema()+".[OCP_Contracts] WHERE NOT [CLIENT] IS NULL GROUP BY [CLIENT]";
        public Clients()
        {
            /*String[] clientsArr = clients.Split(new char[] {';'});
            foreach (String client in clientsArr)
            {
                Add(client);
            }*/
            new idMaker(titleSQL, titleIDs);
            foreach (object client in titleIDs)
            {
                Add(client.ToString());
            }
        }
    }


}

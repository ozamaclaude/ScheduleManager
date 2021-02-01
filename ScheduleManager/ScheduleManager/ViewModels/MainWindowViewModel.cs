using Microsoft.Office.Interop.Outlook;
using Prism.Mvvm;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ScheduleManager.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        private string _title = "Prism Application";
        public string Title
        {
            get { return _title; }
            set { SetProperty(ref _title, value); }
        }

        public MainWindowViewModel()
        {
            GetSchedule2();
        }

        private void GetSchedule()
        {
            Outlook.Application outlook = new Outlook.Application();
            AppointmentItem appointmentItem = outlook.CreateItem(OlItemType.olAppointmentItem);

            if (appointmentItem != null)
            {
                var aaa = "";
                Debug.WriteLine(appointmentItem.Start);
                Debug.WriteLine(appointmentItem.End);
                Debug.WriteLine(appointmentItem.Subject);
            }
        }
        private void GetSchedule2()
        {
            Outlook.Application outlook = new Outlook.Application();
            NameSpace ns = outlook.GetNamespace("MAPI");
            MAPIFolder oFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            
            Items oItems = oFolder.Items;
            AppointmentItem oAppoint = oItems.GetFirst();
            while (oAppoint != null)
            {
                Debug.WriteLine(oAppoint.Subject);
                Debug.WriteLine(oAppoint.Start);
                Debug.WriteLine(oAppoint.End);

                oAppoint = oItems.GetNext();
            }
        }
    }
}

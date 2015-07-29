using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace OCP_2._0
{
    /// <summary>
    /// Interaction logic for NewContract.xaml
    /// </summary>
    public partial class NewContract : Window
    {
        Contract contract = new Contract();
        public NewContract()
        {
            InitializeComponent();
            setBindings();
        }
        private void setBindings()
        {
            TypesofAgreement typesofAgreement = new TypesofAgreement();
            Binding binding1 = new Binding();

            binding1.Source = typesofAgreement;
            cboTypeofAgreement.SetBinding(ListBox.ItemsSourceProperty, binding1);

            Statuses statuses = new Statuses();
            Binding binding2 = new Binding();

            binding2.Source = statuses;
            cboStatus.SetBinding(ListBox.ItemsSourceProperty, binding2);

            UIDepartments departments = new UIDepartments();
            Binding binding3 = new Binding();

            binding3.Source = departments;
            cboUIDepartment.SetBinding(ListBox.ItemsSourceProperty, binding3);


            ReimbursementPeriods reimbursementPeriods = new ReimbursementPeriods();
            Binding binding4 = new Binding();

            binding4.Source = reimbursementPeriods;
            cboReimbursement1Period.SetBinding(ListBox.ItemsSourceProperty, binding4);
            cboReimbursement2Period.SetBinding(ListBox.ItemsSourceProperty, binding4);
            cboReimbursement3Period.SetBinding(ListBox.ItemsSourceProperty, binding4);
            cboTravelReimbursementPeriod.SetBinding(ListBox.ItemsSourceProperty, binding4);
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            contract.addNew(txtClient.Text, cboStatus.Text, txtPeopleSoftContractNum.Text, cboTypeofAgreement.Text, cboUIDepartment.Text, txtSpecialty.Text,
                txtDurationofContract.Text, txtSummaryofServices.Text, txtDepartmentContact.Text, txtServiceClientLocation.Text, dtHardTermination.Text, dtNextAutoRenewal.Text,
                dtEffective.Text, dtCommitteeApproval.Text, txtFrequencyofServices.Text, txtCredentialingProcess.Text, dtMostRecentAnnualWorksheet.Text, chkBillingbyUI.IsChecked,
                chkBAANeededCompleted.IsChecked, whatConsultorProcedure(), txtReimbursement1.Text, txtReimbursement2.Text, txtReimbursement3.Text, txtTravelReimbursement.Text,
                cboReimbursement1Period.Text, cboReimbursement2Period.Text, cboReimbursement3Period.Text, cboTravelReimbursementPeriod.Text, txtIndividualResponsible.Text,
                txtFMR.Text, dtLastFMRReview.Text, txtMFK.Text, txtTimesheetInformation.Text, txtEstimatedAnnualValue.Text, txtReimbursementNotes.Text, txtProviders.Text);
        }
        private int whatConsultorProcedure()
        {
            if (rdoConsultative.IsChecked == true)
            {
                return 1;
            }
            else if (rdoBoth.IsChecked == true)
            {
                return 3;
            }
            else if (rdoProcedural.IsChecked == true)
            {
                return 2;
            }
            else if (rdoNeither.IsChecked == true)
            {
                return 4;
            }
            return 0;
        }

        private void btnRevert_Click(object sender, RoutedEventArgs e)
        {
            txtClient.Text = null;
            txtCredentialingProcess.Text = null;
            txtDepartmentContact.Text = null;
            txtDurationofContract.Text = null;
            txtEstimatedAnnualValue.Text = null;
            txtFMR.Text = null;
            txtFrequencyofServices.Text = null;
            txtIndividualResponsible.Text = null;
            txtMFK.Text = null;
            txtPeopleSoftContractNum.Text = null;
            txtProviders.Text = null;
            txtReimbursement1.Text = null;
            txtReimbursement2.Text = null;
            txtReimbursement3.Text = null;
            txtReimbursementNotes.Text = null;
            txtServiceClientLocation.Text = null;
            txtSpecialty.Text = null;
            txtSummaryofServices.Text = null;
            txtTimesheetInformation.Text = null;
            txtTravelReimbursement.Text = null;

            dtCommitteeApproval.Text = null;
            dtEffective.Text = null;
            dtHardTermination.Text = null;
            dtLastFMRReview.Text = null;
            dtMostRecentAnnualWorksheet.Text = null;
            dtNextAutoRenewal.Text = null;

            cboReimbursement1Period.SelectedIndex = -1;
            cboReimbursement2Period.SelectedIndex = -1;
            cboReimbursement3Period.SelectedIndex = -1;
            cboStatus.SelectedIndex = -1;
            cboTravelReimbursementPeriod.SelectedIndex = -1;
            cboTypeofAgreement.SelectedIndex = -1;
            cboUIDepartment.SelectedIndex = -1;

            rdoBoth.IsChecked = false;
            rdoConsultative.IsChecked = false;
            rdoNeither.IsChecked = false;
            rdoProcedural.IsChecked = false;

            chkBAANeededCompleted.IsChecked = false;
            chkBillingbyUI.IsChecked = false;
        }
    }
}

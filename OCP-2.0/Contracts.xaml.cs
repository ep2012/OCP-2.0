using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
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
    /// Interaction logic for Contracts.xaml
    /// </summary>
    /// 

    public partial class Contracts : Window
    {
        Contract contract;

        enum ButtonPressed {Consultative, Both, Procedural, Neither, Null};

        ButtonPressed buttonPressed;
        public Contracts()
        {
            InitializeComponent();

            contract = new Contract();
            setBindings();
            setWindow();
            hideAllImages();            
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
            lstStatusSelect.SetBinding(ListBox.ItemsSourceProperty, binding2);

            UIDepartments departments = new UIDepartments();
            Binding binding3 = new Binding();

            binding3.Source = departments;
            cboUIDepartment.SetBinding(ListBox.ItemsSourceProperty, binding3);
            cboUIDepartmentSearch.SetBinding(ListBox.ItemsSourceProperty, binding3);

            ReimbursementPeriods reimbursementPeriods = new ReimbursementPeriods();
            Binding binding4 = new Binding();

            binding4.Source = reimbursementPeriods;
            cboReimbursement1Period.SetBinding(ListBox.ItemsSourceProperty, binding4);
            cboReimbursement2Period.SetBinding(ListBox.ItemsSourceProperty, binding4);
            cboReimbursement3Period.SetBinding(ListBox.ItemsSourceProperty, binding4);
            cboTravelReimbursementPeriod.SetBinding(ListBox.ItemsSourceProperty, binding4);

            Clients clients = new Clients();
            Binding binding5 = new Binding();

            binding5.Source = clients;
            cboClients.SetBinding(ListBox.ItemsSourceProperty, binding5);

        }
        private void refreshFilterOptions(String statusSQL)
        {
            Statuses statuses = new Statuses();
            statuses.resetStatuses(statusSQL);
            Binding binding2 = new Binding();

            binding2.Source = statuses;
            lstStatusSelect.SetBinding(ListBox.ItemsSourceProperty, binding2);
        }

        private void setRadioButtons(String number)
        {
            if (number == "1")
            {
                rdoConsultative.IsChecked = true;
            }
            else if (number == "3")
            {
                rdoBoth.IsChecked = true;
            }
            else if (number == "2")
            {
                rdoProcedural.IsChecked = true;
            }
            else if (number == "4")
            {
                rdoNeither.IsChecked = true;
            }
            else
            {
                rdoConsultative.IsChecked = false;

                rdoBoth.IsChecked = false;

                rdoProcedural.IsChecked = false;

                rdoNeither.IsChecked = false;
            }
        }

        private void btnPreviousRecord_Click(object sender, RoutedEventArgs e)
        {
            contract.decrement();
            setWindow();
            hideAllImages();
        }

        private void btnNextRecord_Click(object sender, RoutedEventArgs e)
        {
            contract.increment();
            setWindow();
            hideAllImages();
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            cboClients.SelectedIndex = -1;
            cboUIDepartmentSearch.SelectedIndex = -1;
            chkActive.IsChecked = false;
            chkInactive.IsChecked = false;
            chkInProgress.IsChecked = false;
            chkLessThanNinety.IsChecked = false;

            setBindings();
            contract.filter("", "", new String[0], false);
            setWindow();
            hideAllImages();
        }

        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            NewContract win = new NewContract();
            win.Left = this.Left;
            win.Top = this.Top;
            win.Show();
        }

        private void rdo_Checked(object sender, RoutedEventArgs e)
        {
            if ((contract.getValue(10) == "1" && rdoConsultative.IsChecked == false) ||
                (contract.getValue(10) == "2" && rdoBoth.IsChecked == false) ||
                (contract.getValue(10) == "3" && rdoProcedural.IsChecked == false) ||
                (contract.getValue(10) == "4" && rdoNeither.IsChecked == false))
            {
                imgTypeofService.Visibility = Visibility.Visible;
            }
            else
            {

                imgTypeofService.Visibility = Visibility.Hidden;
            }
        }
        private void hideAllImages()
        {
            imgBAANeededCompleted.Visibility = Visibility.Hidden;
            imgBillingbyUI.Visibility = Visibility.Hidden;
            imgClient.Visibility = Visibility.Hidden;
            imgCommitteeApprovalDate.Visibility = Visibility.Hidden;
            imgCredentialingProcess.Visibility = Visibility.Hidden;
            imgDepartmentContact.Visibility = Visibility.Hidden;
            imgDurationofContract.Visibility = Visibility.Hidden;
            imgEffectiveDate.Visibility = Visibility.Hidden;
            imgFrequencyofServices.Visibility = Visibility.Hidden;
            imgHardTerminationDate.Visibility = Visibility.Hidden;
            imgMostRecentAnnualWorksheetDate.Visibility = Visibility.Hidden;
            imgNextAutoRenewalDate.Visibility = Visibility.Hidden;
            imgPeopleSoftContractNumber.Visibility = Visibility.Hidden;
            imgServiceClientLocation.Visibility = Visibility.Hidden;
            imgSpecialty.Visibility = Visibility.Hidden;
            imgStatus.Visibility = Visibility.Hidden;
            imgSummaryofServices.Visibility = Visibility.Hidden;
            imgTypeofAgreement.Visibility = Visibility.Hidden;
            imgTypeofService.Visibility = Visibility.Hidden;
            imgUIDepartment.Visibility = Visibility.Hidden;

            img2EstimatedAnnualValue.Visibility = Visibility.Hidden;
            img2FMRSources.Visibility = Visibility.Hidden;
            img2IndividualResponsibleforPaymentCompliance.Visibility = Visibility.Hidden;
            img2LastFMRReviewDate.Visibility = Visibility.Hidden;
            img2MFKtobeused.Visibility = Visibility.Hidden;
            img2Notes.Visibility = Visibility.Hidden;
            img2Reimbursement1.Visibility = Visibility.Hidden;
            img2Reimbursement2.Visibility = Visibility.Hidden;
            img2Reimbursement3.Visibility = Visibility.Hidden;
            img2TimesheetInformation.Visibility = Visibility.Hidden;
            img2TravelReimbursement.Visibility = Visibility.Hidden;

            img3Providers.Visibility = Visibility.Hidden;
        }

        private void rdoConsultative_Click(object sender, RoutedEventArgs e)
        {
            if (buttonPressed == ButtonPressed.Consultative)
            {
                rdoConsultative.IsChecked = false;
                buttonPressed = ButtonPressed.Null;
            }
            else
            {
                buttonPressed = ButtonPressed.Consultative;
            }
        }

        private void rdoProcedural_Click(object sender, RoutedEventArgs e)
        {
            if (buttonPressed == ButtonPressed.Procedural)
            {
                rdoProcedural.IsChecked = false;
                buttonPressed = ButtonPressed.Null;
            }
            else
            {
                buttonPressed = ButtonPressed.Procedural;
            }
        }

        private void rdoNeither_Click(object sender, RoutedEventArgs e)
        {
            if (buttonPressed == ButtonPressed.Neither)
            {
                rdoNeither.IsChecked = false;
                buttonPressed = ButtonPressed.Null;
            }
            else
            {
                buttonPressed = ButtonPressed.Neither;
            }
        }

        private void rdoBoth_Click(object sender, RoutedEventArgs e)
        {
            if (buttonPressed == ButtonPressed.Both)
            {
                rdoBoth.IsChecked = false;
                buttonPressed = ButtonPressed.Null;
            }
            else
            {
                buttonPressed = ButtonPressed.Both;
            }
        }

        private void txtClient_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtClient, contract.getValue(2), imgClient);
        }

        private void btnXStatus_Click(object sender, RoutedEventArgs e)
        {
            cboStatus.SelectedIndex = -1;
            if ("" != contract.getValue(3))
            {
                imgStatus.Visibility = Visibility.Visible;
            }
            else
            {
                imgStatus.Visibility = Visibility.Hidden;
            }
        }

        private void btnXTypeofAgreement_Click(object sender, RoutedEventArgs e)
        {
            cboTypeofAgreement.SelectedIndex = -1;
            if ("" != contract.getValue(4))
            {
                imgTypeofAgreement.Visibility = Visibility.Visible;
            }
            else
            {
                imgTypeofAgreement.Visibility = Visibility.Hidden;
            }
        }

        private void btnXUIDepartment_Click(object sender, RoutedEventArgs e)
        {
            cboUIDepartment.SelectedIndex = -1;

            if ("" != contract.getValue(5))
            {
                imgUIDepartment.Visibility = Visibility.Visible;
            }
            else
            {
                imgUIDepartment.Visibility = Visibility.Hidden;
            }
        }

        private void cboStatus_LostFocus(object sender, RoutedEventArgs e)
        {
            setImageVisibility(cboStatus, contract.getValue(3), imgStatus);
        }

        private void txtProviders_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtProviders, contract.getValue(37), img3Providers);
        }

        private void txtReimbursementNotes_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtReimbursementNotes, contract.getValue(31), img2Notes);
        }

        private void cboTypeofAgreement_LostFocus(object sender, RoutedEventArgs e)
        {
            setImageVisibility(cboTypeofAgreement, contract.getValue(4), imgTypeofAgreement);
        }

        private void cboUIDepartment_LostFocus(object sender, RoutedEventArgs e)
        {
            setImageVisibility(cboUIDepartment, contract.getValue(5), imgUIDepartment);
        }

        private void txtSpecialty_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtSpecialty, contract.getValue(6), imgSpecialty);
        }

        private void txtDepartmentContact_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtDepartmentContact, contract.getValue(7), imgDepartmentContact);
        }

        private void txtServiceClientLocation_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtServiceClientLocation, contract.getValue(8), imgServiceClientLocation);
        }

        private void setImageVisibility(TextBox textbox, String contractValue, Image image)
        {
            if (textbox.Text != contractValue)
            {
                image.Visibility = Visibility.Visible;
            }
            else
            {
                image.Visibility = Visibility.Hidden;
            }
        }
        private void setImageVisibility(ComboBox textbox, String contractValue, Image image)
        {
            if (textbox.Text != contractValue)
            {
                image.Visibility = Visibility.Visible;
            }
            else
            {
                image.Visibility = Visibility.Hidden;
            }
        }
        private void setImageVisibility(DatePicker textbox, String contractValue, Image image)
        {
            try
            {
                if (!DateTime.Parse(textbox.Text).Equals(DateTime.Parse(contractValue)))
                {
                    image.Visibility = Visibility.Visible;
                }
                else
                {
                    image.Visibility = Visibility.Hidden;
                }
            }
            catch(Exception)
            {
                if (textbox.Text != contractValue)
                {
                    image.Visibility = Visibility.Visible;
                }
                else
                {
                    image.Visibility = Visibility.Hidden;
                }
            }
        }

        private void setWindow()
        {
            buttonPressed = ButtonPressed.Null;

            lblRecordCount.Content = "Record " + (contract.getIndex() + 1) + "/" + contract.getLength();

            DateTime date = new DateTime();
            Boolean bltsend = new Boolean();
            if (DateTime.TryParse(contract.getValue(1), out date))
                lblDateRevised.Content = date.ToShortDateString();
            else
                lblDateRevised.Content = null;
            txtClient.Text = contract.getValue(2);
            cboStatus.Text = contract.getValue(3);
            cboTypeofAgreement.Text = contract.getValue(4);
            cboUIDepartment.Text = contract.getValue(5);
            txtSpecialty.Text = contract.getValue(6);
            txtDepartmentContact.Text = contract.getValue(7);
            txtServiceClientLocation.Text = contract.getValue(8);
            txtSummaryofServices.Text = contract.getValue(9);
            setRadioButtons(contract.getValue(10));
            if (DateTime.TryParse(contract.getValue(11), out date))
                dtNextAutoRenewal.Text = date.ToShortDateString();
            else
                dtNextAutoRenewal.Text = null;
            if (DateTime.TryParse(contract.getValue(12), out date))
                dtHardTermination.Text = date.ToShortDateString();
            else
                dtHardTermination.Text = null;
            if (DateTime.TryParse(contract.getValue(13), out date))
                dtEffective.Text = date.ToShortDateString();
            else
                dtEffective.Text = null;
            if (DateTime.TryParse(contract.getValue(14), out date))
                dtCommitteeApproval.Text = date.ToShortDateString();
            else
                dtCommitteeApproval.Text = null;
            txtDurationofContract.Text = contract.getValue(15);
            txtFrequencyofServices.Text = contract.getValue(16);
            txtReimbursement1.Text = contract.getValue(17);
            txtReimbursement2.Text = contract.getValue(18);
            txtReimbursement3.Text = contract.getValue(19);
            txtTravelReimbursement.Text = contract.getValue(20);
            cboReimbursement1Period.Text = contract.getValue(21);
            cboReimbursement2Period.Text = contract.getValue(22);
            cboReimbursement3Period.Text = contract.getValue(23);
            cboTravelReimbursementPeriod.Text = contract.getValue(24);
            txtIndividualResponsible.Text = contract.getValue(25);
            txtFMR.Text = contract.getValue(26);
            if (DateTime.TryParse(contract.getValue(27), out date))
                dtLastFMRReview.Text = date.ToShortDateString();
            else
                dtLastFMRReview.Text = null;
            txtMFK.Text = contract.getValue(28);
            txtTimesheetInformation.Text = contract.getValue(29);
            txtEstimatedAnnualValue.Text = contract.getValue(30);
            txtReimbursementNotes.Text = contract.getValue(31);
            txtPeopleSoftContractNum.Text = contract.getValue(32);
            txtCredentialingProcess.Text = contract.getValue(33);
            if (DateTime.TryParse(contract.getValue(34), out date))
                dtMostRecentAnnualWorksheet.Text = date.ToShortDateString();
            else
                dtMostRecentAnnualWorksheet.Text = null;
            if (Boolean.TryParse(contract.getValue(35), out bltsend))
                chkBillingbyUI.IsChecked = bltsend;
            else
                chkBillingbyUI.IsChecked = false;
            if (Boolean.TryParse(contract.getValue(36), out bltsend))
                chkBAANeededCompleted.IsChecked = bltsend;
            else
                chkBAANeededCompleted.IsChecked = false;
            txtProviders.Text = contract.getValue(37);

            updateExpirationCountdown();
        }

        private void updateExpirationCountdown()
        {
            if (contract.getValue(3) != null)
            {
                if (!contract.getValue(3).Contains("Inactive"))
                {
                    DateTime nextAutoRenewal;
                    if (DateTime.TryParse(dtNextAutoRenewal.Text, out nextAutoRenewal))
                    {
                        int expiration = (int) nextAutoRenewal.Subtract(DateTime.Today).TotalDays;
                        if (expiration < 0)
                        {
                            lblExpirationCountdown.Content = "The date has passed";
                        }
                        else
                        {
                            lblExpirationCountdown.Content = expiration + " days";
                        }
                    }
                    else
                    {
                        lblExpirationCountdown.Content = "The date is not set.";

                    }
                }
                else
                {
                    lblExpirationCountdown.Content = "This contract is inactive";
                }
            }
            else
            {
                lblExpirationCountdown.Content = "This contract is inactive";
            }
        }

        private void txtSummaryofServices_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtSummaryofServices, contract.getValue(9), imgSummaryofServices);
        }

        private void txtPeopleSoftContractNum_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtPeopleSoftContractNum, contract.getValue(32), imgPeopleSoftContractNumber);
        }

        private void txtDurationofContract_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtDurationofContract, contract.getValue(15), imgDurationofContract);
        }

        private void txtFrequencyofServices_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtFrequencyofServices, contract.getValue(16), imgFrequencyofServices);
        }

        private void txtCredentialingProcess_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtCredentialingProcess, contract.getValue(33), imgCredentialingProcess);
        }

        private void txtEstimatedAnnualValue_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtEstimatedAnnualValue, contract.getValue(30), img2EstimatedAnnualValue);
        }

        private void txtIndividualResponsible_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtIndividualResponsible, contract.getValue(25), img2IndividualResponsibleforPaymentCompliance);
        }

        private void txtTimesheetInformation_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtTimesheetInformation, contract.getValue(29), img2TimesheetInformation);
        }

        private void txtFMR_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtFMR, contract.getValue(26), img2FMRSources);
        }

        private void txtMFK_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtMFK, contract.getValue(28), img2MFKtobeused);
        }

        private void dtNextAutoRenewal_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            setImageVisibility(dtNextAutoRenewal, contract.getValue(11), imgNextAutoRenewalDate);
        }

        private void dtHardTermination_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            setImageVisibility(dtHardTermination, contract.getValue(12), imgHardTerminationDate);
        }

        private void dtCommitteeApproval_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            setImageVisibility(dtCommitteeApproval, contract.getValue(14), imgCommitteeApprovalDate);
        }

        private void dtEffective_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            setImageVisibility(dtEffective, contract.getValue(13), imgEffectiveDate);
        }

        private void dtMostRecentAnnualWorksheet_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            setImageVisibility(dtMostRecentAnnualWorksheet, contract.getValue(34), imgMostRecentAnnualWorksheetDate);
        }

        private void chkBillingbyUI_Click(object sender, RoutedEventArgs e)
        {
            if (chkBillingbyUI.IsChecked != Boolean.Parse(contract.getValue(35)))
            {
                imgBillingbyUI.Visibility = Visibility.Visible;
            }
            else
            {
                imgBillingbyUI.Visibility = Visibility.Hidden;
            }
        }

        private void chkBAANeededCompleted_Click(object sender, RoutedEventArgs e)
        {
            if (chkBAANeededCompleted.IsChecked != Boolean.Parse(contract.getValue(36)))
            {
                imgBAANeededCompleted.Visibility = Visibility.Visible;
            }
            else
            {
                imgBAANeededCompleted.Visibility = Visibility.Hidden;
            }
        }

        private void dtLastFMRReview_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            setImageVisibility(dtLastFMRReview, contract.getValue(27), img2LastFMRReviewDate);
        }

        private void txtReimbursement1_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtReimbursement1, contract.getValue(17), img2Reimbursement1);
            if (cboReimbursement1Period.Text != contract.getValue(21))
            {
                img2Reimbursement1.Visibility = Visibility.Visible;
            }
        }

        private void txtReimbursement2_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtReimbursement2, contract.getValue(18), img2Reimbursement2);
            if (cboReimbursement2Period.Text != contract.getValue(22))
            {
                img2Reimbursement2.Visibility = Visibility.Visible;
            }
        }

        private void txtReimbursement3_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtReimbursement3, contract.getValue(19), img2Reimbursement3);
            if (cboReimbursement3Period.Text != contract.getValue(23))
            {
                img2Reimbursement3.Visibility = Visibility.Visible;
            }
        }

        private void txtTravelReimbursement_SelectionChanged(object sender, RoutedEventArgs e)
        {
            setImageVisibility(txtTravelReimbursement, contract.getValue(20), img2TravelReimbursement);
            if (cboTravelReimbursementPeriod.Text != contract.getValue(24))
            {
                img2TravelReimbursement.Visibility = Visibility.Visible;
            }
        }

        private void cboReimbursement1Period_LostFocus(object sender, RoutedEventArgs e)
        {
            setImageVisibility(cboReimbursement1Period, contract.getValue(21), img2Reimbursement1);
            if (txtReimbursement1.Text != contract.getValue(17))
            {
                img2Reimbursement1.Visibility = Visibility.Visible;
            }
        }

        private void cboReimbursement2Period_LostFocus(object sender, RoutedEventArgs e)
        {
            setImageVisibility(cboReimbursement2Period, contract.getValue(22), img2Reimbursement2);
            if (txtReimbursement2.Text != contract.getValue(18))
            {
                img2Reimbursement2.Visibility = Visibility.Visible;
            }
        }

        private void cboReimbursement3Period_LostFocus(object sender, RoutedEventArgs e)
        {
            setImageVisibility(cboReimbursement3Period, contract.getValue(23), img2Reimbursement3);
            if (txtReimbursement3.Text != contract.getValue(19))
            {
                img2Reimbursement3.Visibility = Visibility.Visible;
            }
        }

        private void cboTravelReimbursementPeriod_LostFocus(object sender, RoutedEventArgs e)
        {
            setImageVisibility(cboTravelReimbursementPeriod, contract.getValue(24), img2TravelReimbursement);
            if (txtTravelReimbursement.Text != contract.getValue(20))
            {
                img2TravelReimbursement.Visibility = Visibility.Visible;
            }
        }

        private void btnRevert_Click(object sender, RoutedEventArgs e)
        {
            setWindow();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            contract.updateSelection(txtClient.Text, cboStatus.Text, txtPeopleSoftContractNum.Text, cboTypeofAgreement.Text, cboUIDepartment.Text, txtSpecialty.Text,
                txtDurationofContract.Text, txtSummaryofServices.Text, txtDepartmentContact.Text, txtServiceClientLocation.Text, dtHardTermination.Text, dtNextAutoRenewal.Text,
                dtEffective.Text, dtCommitteeApproval.Text, txtFrequencyofServices.Text, txtCredentialingProcess.Text, dtMostRecentAnnualWorksheet.Text, chkBillingbyUI.IsChecked,
                chkBAANeededCompleted.IsChecked, whatConsultorProcedure(), txtReimbursement1.Text, txtReimbursement2.Text, txtReimbursement3.Text, txtTravelReimbursement.Text,
                cboReimbursement1Period.Text, cboReimbursement2Period.Text, cboReimbursement3Period.Text, cboTravelReimbursementPeriod.Text, txtIndividualResponsible.Text,
                txtFMR.Text, dtLastFMRReview.Text, txtMFK.Text, txtTimesheetInformation.Text, txtEstimatedAnnualValue.Text, txtReimbursementNotes.Text, txtProviders.Text);
            contract.setValues();
            setWindow();
            hideAllImages();
        }
        private int whatConsultorProcedure()
        {
            if (rdoConsultative.IsChecked == true)
            {
                return 1;
            }
            else if (rdoBoth.IsChecked == true)
            {
                return 2;
            }
            else if (rdoProcedural.IsChecked == true)
            {
                return 3;
            }
            else if (rdoNeither.IsChecked == true)
            {
                return 4;
            }
            return 0;
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            contract.deleteSelection();
            setWindow();
        }

        private void filter()
        {
            String[] sActionItemsSelected = new String[lstStatusSelect.SelectedItems.Count];
            for (int x = 0; x < lstStatusSelect.SelectedItems.Count; ++x)
            {
                lstStatusSelect.SelectedItems.CopyTo(sActionItemsSelected, 0);
            }
            refreshFilterOptions(contract.filter(cboClients.Text, cboUIDepartmentSearch.Text, sActionItemsSelected, Boolean.Parse(chkLessThanNinety.IsChecked.ToString())));
            setWindow();
            hideAllImages();
        }

        private void btnExportSubset_Click(object sender, RoutedEventArgs e)
        {
            String sqlString = "SELECT * FROM [REVINT]." + contract.getSchema() + ".[OCP_Contracts] WHERE 1 = 1";

            if (cboClients.Text != "")
            {
                sqlString += " AND CLIENT = '" + cboClients.Text + "'";
            }
            if (cboUIDepartmentSearch.Text != "")
            {
                sqlString += " AND UI_DEPARTMENT = '" + cboUIDepartmentSearch.Text + "'";
            }

            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";

            //create an OdbcConnection object and connect it to the data source.
            using (OdbcConnection dbConnection = new OdbcConnection(cxnString))
            {
                //open OdbcConnection object
                dbConnection.Open();

                //Create adapter from connection and sql to obtain desired data
                OdbcDataAdapter dadapter = new OdbcDataAdapter(sqlString, dbConnection);

                //Create a table and fill it with the data from the adapter
                DataTable dtable = new DataTable();
                dadapter.Fill(dtable);

                //set the contents of the gui grid table to the data table created
                this.dta.ItemsSource = dtable.DefaultView;
                this.dta.CanUserAddRows = false;

                //Close connection
                dbConnection.Close();
            }
            export(dta);
        }

            
        private void BorderAround(Microsoft.Office.Interop.Excel.Range range, int color)
        {
            Microsoft.Office.Interop.Excel.Borders borders = range.Borders;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders.Color = color;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders = null;
        }

        private void btnExportAll_Click(object sender, RoutedEventArgs e)
        {
            String sqlString = "SELECT * FROM [REVINT]." + contract.getSchema() + ".[OCP_Contracts]";

            String cxnString = "Driver={SQL Server};Server=HC-sql7;Database=REVINT;Trusted_Connection=yes;";

            //create an OdbcConnection object and connect it to the data source.
            using (OdbcConnection dbConnection = new OdbcConnection(cxnString))
            {
                //open OdbcConnection object
                dbConnection.Open();

                //Create adapter from connection and sql to obtain desired data
                OdbcDataAdapter dadapter = new OdbcDataAdapter(sqlString, dbConnection);

                //Create a table and fill it with the data from the adapter
                DataTable dtable = new DataTable();
                dadapter.Fill(dtable);

                //set the contents of the gui grid table to the data table created
                this.dta.ItemsSource = dtable.DefaultView;
                this.dta.CanUserAddRows = false;

                //Close connection
                dbConnection.Close();
            }
            export(dta);
        }

        public void export(DataGrid dta)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = new Microsoft.Office.Interop.Excel.Worksheet();
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;

            try
            {
                // Put Column Header into excel work sheet
                for (int i = 0; i < dta.Columns.Count; i++)
                {
                    worksheet.Range["A1"].Offset[0, i].Value = dta.Columns[i].Header;
                }

                Microsoft.Office.Interop.Excel.Range firstRow = (Microsoft.Office.Interop.Excel.Range)worksheet.Rows[1];
                firstRow.Cells.Interior.ColorIndex = 36;

                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                firstRow.EntireRow.Font.Bold = true;

                BorderAround(firstRow, 0);

                for (int rowIndex = 0; rowIndex < dta.Items.Count; rowIndex++)
                {
                    for (int columnIndex = 0; columnIndex < dta.Columns.Count; columnIndex++)
                    {
                        worksheet.Range["A2"].Offset[rowIndex, columnIndex].Value =
                          dtaToString((dta.Items[rowIndex] as DataRowView).Row.ItemArray[columnIndex]);
                    }
                }
                worksheet.Columns.AutoFit();
                app.Visible = true;
            }
            catch (Exception)
            {
                Console.Write("Error");
            }
        }

        private String dtaToString(object data)
        {
            if (data != null)
            {
                return data.ToString();
            }
            else
            {
                return "";
            }
        }

        private void btnFirst_Click(object sender, RoutedEventArgs e)
        {
            contract.first();
            setWindow();
            hideAllImages();
        }

        private void btnFilter_Click(object sender, RoutedEventArgs e)
        {
            filter();
        }

        private void btnDuplicate_Click(object sender, RoutedEventArgs e)
        {
            NewContract win = new NewContract();

            win.txtClient.Text = txtClient.Text;
            win.txtCredentialingProcess.Text = txtCredentialingProcess.Text;
            win.txtDepartmentContact.Text = txtDepartmentContact.Text;
            win.txtDurationofContract.Text = txtDurationofContract.Text;
            win.txtEstimatedAnnualValue.Text = txtEstimatedAnnualValue.Text;
            win.txtFMR.Text = txtFMR.Text;
            win.txtFrequencyofServices.Text = txtFrequencyofServices.Text;
            win.txtIndividualResponsible.Text = txtIndividualResponsible.Text;
            win.txtMFK.Text = txtMFK.Text;
            win.txtPeopleSoftContractNum.Text = txtPeopleSoftContractNum.Text;
            win.txtProviders.Text = txtProviders.Text;
            win.txtReimbursement1.Text = txtReimbursement1.Text;
            win.txtReimbursement2.Text = txtReimbursement2.Text;
            win.txtReimbursement3.Text = txtReimbursement3.Text;
            win.txtReimbursementNotes.Text = txtReimbursementNotes.Text;
            win.txtServiceClientLocation.Text = txtServiceClientLocation.Text;
            win.txtSpecialty.Text = txtSpecialty.Text;
            win.txtSummaryofServices.Text = txtSummaryofServices.Text;
            win.txtTimesheetInformation.Text = txtTimesheetInformation.Text;
            win.txtTravelReimbursement.Text = txtTravelReimbursement.Text;

            win.dtCommitteeApproval.Text = dtCommitteeApproval.Text;
            win.dtEffective.Text = dtEffective.Text;
            win.dtHardTermination.Text = dtHardTermination.Text;
            win.dtLastFMRReview.Text = dtLastFMRReview.Text;
            win.dtMostRecentAnnualWorksheet.Text = dtMostRecentAnnualWorksheet.Text;
            win.dtNextAutoRenewal.Text = dtNextAutoRenewal.Text;

            win.cboReimbursement1Period.Text = cboReimbursement1Period.Text;
            win.cboReimbursement2Period.Text = cboReimbursement2Period.Text;
            win.cboReimbursement3Period.Text = cboReimbursement3Period.Text;
            win.cboStatus.Text = cboStatus.Text;
            win.cboTravelReimbursementPeriod.Text = cboTravelReimbursementPeriod.Text;
            win.cboTypeofAgreement.Text = cboTypeofAgreement.Text;
            win.cboUIDepartment.Text = cboUIDepartment.Text;
            
            win.chkBAANeededCompleted.IsChecked = chkBAANeededCompleted.IsChecked;
            win.chkBillingbyUI.IsChecked = chkBillingbyUI.IsChecked;

            win.rdoBoth.IsChecked = rdoBoth.IsChecked;
            win.rdoConsultative.IsChecked = rdoConsultative.IsChecked;
            win.rdoNeither.IsChecked = rdoNeither.IsChecked;
            win.rdoProcedural.IsChecked = rdoProcedural.IsChecked;

            win.Left = this.Left;
            win.Top = this.Top;
            win.Show();
        }
    }
}
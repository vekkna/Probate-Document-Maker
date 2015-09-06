/* This app creates the skeleton of some of the more often used probate documents.
 * It creates a windows form with a check box for each document it can create, and 
 * textboxes where the user can input the information that those docuements need.
 * Checking a checkbox activated the fields needed by that document and multiple
    documents can be created at once.*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Novacode;

namespace Probate_Document_Maker {
    public partial class Form1 : Form {

        // These dictionaries count how many documents currently require particular input fields to be filled in.
        // Any fields that are needed by at least one document will be activated below.
        // Any fields that are not needed by any documents will be deactivated.
        Dictionary<TextBox, int> fieldRequirements;
        Dictionary<Panel, int> radioButtonRequirements;
        bool needsCA24;

        public Form1() {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) {
            // Add each textbox in the Input Fields panel to the fieldRequirements list and set the 
            // number of documents requiring it to zero.
            fieldRequirements = new Dictionary<TextBox, int>();
            foreach (TextBox t in inputFields.Controls.OfType<TextBox>()) {
                fieldRequirements.Add(t, 0);
                t.Enabled = false;
            }
            // Add each radiobutton group in the Input Fields panel to the fieldRequirements list and set the 
            // number of documents requiring it to zero.
            radioButtonRequirements = new Dictionary<Panel, int>();
            foreach (Panel p in inputFields.Controls.OfType<Panel>()) {
                radioButtonRequirements.Add(p, 0);
                p.Enabled = false;
            }
        }
        // This is called whenever the user ticks a document checkbox, and is passed an array of fields that that document requires.
        void UpdateFieldRequirements(TextBox[] fields, CheckBox checkbox) {
            foreach (TextBox t in fields) {
                // If the checkbox is checked, add one to the number of documents requiring those fields
                fieldRequirements[t] += checkbox.Checked ? 1 : -1;
                // Activate any textboxes that are needed by at least one document and deactivate any others.
                t.Enabled = fieldRequirements[t] > 0;
            }
            foreach (Panel p in inputFields.Controls.OfType<Panel>()) {
                /// If the checkbox is unchecked, subtract one from the number of documents requiring those fields
                radioButtonRequirements[p] += checkbox.Checked ? 1 : -1;
                // Activate any textboxes that are needed by at least one document and deactivate any others.
                p.Enabled = radioButtonRequirements[p] > 0;
            }
        }

        // The buttons that generates the documents.
        private void CreateDocuments_Click(object sender, EventArgs e) {
            // For the documents to be safely created, proceed must be set to true.
            bool proceed = false;

            // Check that at least one document checkbox is checked.
            foreach (CheckBox c in listOfDocuments.Controls) {
                if (c.Checked) {
                    proceed = true;
                    break;
                }
            }
            if (!proceed) {
                MessageBox.Show("Please select at least one document to create.", "Warning");
            }
            // If at least one document checkbox is checked,  check that all required input fields have been filled
            else {
                foreach (TextBox t in inputFields.Controls.OfType<TextBox>()) {
                    // if a field is empty and required.
                    if (t.Text.Length == 0 && fieldRequirements[t] > 0) {
                        MessageBox.Show("Please fill all active fields in.", "Warning");
                        proceed = false;
                        break;
                    }
                }
            }
            if (proceed) {
                FolderBrowserDialog browser = new FolderBrowserDialog();
                browser.Description = "Choose where to save the documents";
                if (browser.ShowDialog() == DialogResult.OK) {
                    foreach (CheckBox c in listOfDocuments.Controls) {
                        if (c.Checked && c.Text != "CA 24") {

                            // For each selected document, load from the templates folder the template whose name matches the checkbox text.
                            var doc = DocX.Load(@"..\..\templates\" + c.Text + ".docx");

                            // Scan the template for placeholder text and replace it with the contents of the corresponding textbox
                            doc.ReplaceText("%COURT%", court.Text);
                            doc.ReplaceText("%REGISTRY%", registry.Text);
                            doc.ReplaceText("%DECEASED NAME%", deceasedName.Text);
                            doc.ReplaceText("%DECEASED ADDRESS%", deceasedAddress.Text);
                            doc.ReplaceText("%DECEASED OCCUPATION%", deceasedOccupation.Text);
                            doc.ReplaceText("%DECEASED MARITAL STATUS", maritalStatus.Text);
                            doc.ReplaceText("%PLACE OF DEATH%", placeOfDeath.Text);
                            doc.ReplaceText("%DATE OF DEATH%", dateOfDeath.Text);
                            doc.ReplaceText("%VALUE OF ESTATE%", valueOfEstate.Text);
                            doc.ReplaceText("%DOUBLE VALUE OF ESTATE%", doubleValueOfEstate.Text);
                            doc.ReplaceText("%VALUE OF PERSONAL ESTATE%", valueOfPersonalEstate.Text);
                            doc.ReplaceText("%VALUE OF REAL ESTATE%", valueOfRealEstate.Text);
                            doc.ReplaceText("%HIS/HER%", male.Checked ? "his" : "her");
                            doc.ReplaceText("%HE/SHE%", male.Checked ? "he" : "she");
                            doc.ReplaceText("%TESTATOR/TESTATRIX% ", male.Checked ? "testator" : "testatrix");
                            doc.ReplaceText("%APPLICANT NAME%", applicantName.Text);
                            doc.ReplaceText("%APPLICANT ADDRESS%", applicantAddress.Text);
                            doc.ReplaceText("%APPLICANT OCCUPATION%", applicantOccupation.Text);
                            doc.ReplaceText("%APPLICANT HIM/HER%", applicantMale.Checked ? "him" : "her");
                            doc.ReplaceText("%APPLICANT HIS/HER%", applicantMale.Checked ? "him" : "her");
                            doc.ReplaceText("%APPLICANT RELATION%", relation.Text);
                            doc.ReplaceText("%DATE OF WILL%", dateOfWill.Text);
                            doc.ReplaceText("%OTHER WITNESS NAME%", otherWitnessName.Text);
                            doc.ReplaceText("%OTHER WITNESS OCCUPATION%", otherWitnessOccupation.Text);

                            // Save the template with the deceased name prefixed to the title.
                            doc.SaveAs(@browser.SelectedPath + "\\" + deceasedName.Text + " " + c.Text + ".docx");
                        }
                    }
                }
            }
        }

        // Checkboxes allowing user to select which documents to create. 
        private void oathExecCheck_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[]{court, registry, deceasedName, deceasedAddress, deceasedOccupation, applicantName, placeOfDeath, dateOfDeath,
                applicantAddress, applicantOccupation, relation, valueOfEstate}, sender as CheckBox);
        }

        private void oathAdminCheck_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[]{court, registry, deceasedName, deceasedAddress, deceasedOccupation, applicantName, relation,
            applicantAddress, applicantOccupation, maritalStatus, placeOfDeath, dateOfDeath, valueOfPersonalEstate, valueOfRealEstate}, sender as CheckBox);
        }

        private void adminWillAnnexed_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[] { court, registry, deceasedName, deceasedAddress, deceasedOccupation, applicantName, applicantAddress, applicantOccupation, relation, originalExecutorName, placeOfDeath, dateOfDeath, maritalStatus, valueOfPersonalEstate, valueOfRealEstate }, sender as CheckBox);
        }

        private void adminBond_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[] {court, registry, applicantName, applicantAddress, applicantOccupation, doubleValueOfEstate, relation, deceasedName,
            deceasedAddress, deceasedOccupation }, sender as CheckBox);
        }

        private void mentalCapacitySol_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[] { court, registry, deceasedName, deceasedAddress, applicantName, applicantAddress, applicantOccupation, dateOfWill, otherWitnessName, maritalStatus }, sender as CheckBox);
        }

        private void mentalCapacityDoctor_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[] { court, registry, deceasedName, deceasedAddress, maritalStatus, applicantName, applicantAddress, dateOfWill }, sender as CheckBox);
        }

        private void plightCondition_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[]{court, registry, deceasedName, deceasedAddress, deceasedOccupation, applicantName, applicantAddress, applicantOccupation,
                  dateOfWill, otherWitnessName, otherWitnessOccupation}, sender as CheckBox);
        }

        private void attestingWitness_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[] { court, registry, deceasedName, applicantName, applicantAddress, dateOfWill }, sender as CheckBox);
        }

        private void letterToBank_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[] { deceasedName, deceasedAddress, dateOfDeath }, sender as CheckBox);
        }

        private void affMarketVal_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[] { deceasedName, deceasedAddress, court }, sender as CheckBox);
        }

        private void instructionsProgressCheckbox_CheckedChanged(object sender, EventArgs e) {
            UpdateFieldRequirements(new TextBox[] { deceasedName, deceasedAddress, deceasedOccupation, dateOfDeath, maritalStatus, placeOfDeath, dateOfWill }, sender as CheckBox);
        }

        private void instructionsToolStripMenuItem_Click(object sender, EventArgs e) {
            Instructions instructions = new Instructions();
            instructions.Show();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e) {
            About about = new About();
            about.Show();
        }
    }
}
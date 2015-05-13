using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TXTextControl;
using TXTextControl.DocumentServer.Fields;

namespace tx_custom_fields_SUM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void loadTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadSettings ls = new LoadSettings();
            ls.ApplicationFieldFormat = ApplicationFieldFormat.MSWord;

            textControl1.Load("template.docx", StreamType.WordprocessingML, ls);
        }

        private void mergeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds.ReadXml("sample_db.xml", XmlReadMode.Auto);

            mailMerge1.Merge(ds.Tables[0]);
        }

        private void mailMerge1_FieldMerged(object sender, TXTextControl.DocumentServer.MailMerge.FieldMergedEventArgs e)
        {
            // set the field value to keep field functionality
            // fields that are set in the event are not removed
            e.MergedField = e.MergedField;
        }

        private void mailMerge1_DataRowMerged(object sender, TXTextControl.DocumentServer.MailMerge.DataRowMergedEventArgs e)
        {
            byte[] data = null;

            // create a temporary ServerTextControl to work on the block data
            using (ServerTextControl tx = new ServerTextControl())
            {
                tx.Create();

                // load the block data into the temporary ServerTextControl
                tx.Load(e.MergedRow, BinaryStreamType.InternalUnicodeFormat);

                // loop #1 to find all SUM fields
                foreach (ApplicationField sumField in tx.ApplicationFields)
                {
                    if (sumField.TypeName != "MERGEFIELD")
                        continue;

                    MergeField mf = new MergeField(sumField);

                    // if SUM field found, loop through all fields again total them up
                    if (mf.Name.StartsWith("SUM:") == true)
                    {
                        decimal fSum = 0.0M;

                        // loop #2
                        foreach (ApplicationField appField in tx.ApplicationFields)
                        {
                            if (appField.TypeName != "MERGEFIELD")
                                continue;

                            MergeField mfa = new MergeField(appField);

                            if (mfa.Name == mf.Name.Substring(4))
                            {
                                fSum += Decimal.Parse(mfa.Text,
                                    NumberStyles.AllowCurrencySymbol | 
                                    NumberStyles.AllowDecimalPoint | 
                                    NumberStyles.AllowThousands, 
                                    new CultureInfo("en-US"));
                            }
                        }

                        // set the total number
                        mf.Text = fSum.ToString();
                    }
                }

                // save the complete block to a byte[] array
                tx.Save(out data, BinaryStreamType.InternalUnicodeFormat);
            }

            // load back to the byte[] array to the MailMerge process
            e.MergedRow = data;
        }

        private void mailMergeToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
        {
            mergeToolStripMenuItem.Enabled = (textControl1.ApplicationFields.Count > 0) ? true : false;
        }
    }
}

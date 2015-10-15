/*
Copyright (c) 2015, acknack
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:
* Redistributions of source code must retain the above copyright notice, 
  this list of conditions and the following disclaimer.
* Redistributions in binary form must reproduce the above copyright notice, 
  this list of conditions and the following disclaimer in the documentation 
  and/or other materials provided with the distribution.
* Neither the name of the <organization> nor the　names of its contributors 
  may be used to endorse or promote products derived from this software 
  without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

using System;
using System.Data;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Reflection;

namespace yapdf2pocket
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        SaveFileDialog sfd = new SaveFileDialog();

        public Form1()
        {
            InitializeComponent();

            // Setup comboBox.
            comboBox1.Items.Clear();
            Assembly a = Assembly.LoadFrom(".\\itextsharp.dll");
            foreach (Type t in a.GetExportedTypes())
            {
                if (t.Name == "PageSize")
                {
                    FieldInfo[] f = t.GetFields();
                    foreach (FieldInfo fi in f)
                    {
                        if (fi.FieldType.Name == "Rectangle")
                            comboBox1.Items.Add(fi.Name);
                    }
                }
            }
            var index = comboBox1.FindStringExact("A4");
            comboBox1.SelectedIndex = index;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ofd.Filter = "PDF files(*.pdf)|*.pdf";
            ofd.Title = "Select a PDF file.";
            //ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //CheckPdfProtection(ofd.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ofd.FileName == "")
            {
                MessageBox.Show("PDF file is not selected.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            sfd.Filter = "PDF files(*.pdf)|*.pdf";
            sfd.Title = "Save PocketMod PDF";
            //sfd.RestoreDirectory = true;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                if (ofd.FileName == sfd.FileName)
                {
                    MessageBox.Show("Same PDF file could not be accepted.",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // Processing
                pocketmodstyle();

                ofd.FileName = "";
                sfd.FileName = "";
            }
        }

        private void pocketmodstyle()
        {

            PdfReader pr = new PdfReader(ofd.FileName);
            if (pr.IsEncrypted())
            {
                MessageBox.Show(ofd.FileName + " is protected.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            try
            {

                iTextSharp.text.Rectangle dst = PageSize.GetRectangle(comboBox1.SelectedItem.ToString());

                Document doc = new Document(dst);

                PdfWriter pw = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));

                doc.Open();

                for (int i = 1; i <= pr.NumberOfPages; i++)
                {
                    if (i % 8 == 1) doc.NewPage();

                    PdfContentByte pcb = pw.DirectContent;

                    iTextSharp.text.Rectangle src = pr.GetPageSizeWithRotation(i);

                    int rot = pr.GetPageRotation(i);

                    bool isLandscape = (src.Width > src.Height) ? true : false;

                    float scale;
                    float offset;

                    if ((isLandscape) && (rot == 90 || rot == 270))
                    {
                        scale = dst.Height / (src.Height * 4);
                        offset = (dst.Width / 2 - (scale * src.Width)) / 2;
                    }
                    else
                    {
                        scale = dst.Height / (src.Width * 4);
                        offset = (dst.Width / 2 - (scale * src.Height)) / 2;
                    }

                    PdfImportedPage page;
                    page = pw.GetImportedPage(pr, i);

                    // Affine translation of PocketMod style.
                    if ((rot == 90 || rot == 270) && !isLandscape)
                    {
                        switch (i % 8)
                        {
                            // LEFT SIDE
                            case 1:
                                pcb.AddTemplate(page,
                                    -scale, 0f, 0f, -scale,
                                    (dst.Width / 2) - offset, dst.Height);
                                break;

                            // RIGHT SIDE
                            case 2:
                                pcb.AddTemplate(page,
                                    scale, 0f, 0f, scale,
                                    (dst.Width / 2) + offset, dst.Height * 3 / 4);
                                break;
                            case 3:
                                pcb.AddTemplate(page,
                                    scale, 0f, 0f, scale,
                                    (dst.Width / 2) + offset, dst.Height * 2 / 4);
                                break;
                            case 4:
                                pcb.AddTemplate(page,
                                    scale, 0f, 0f, scale,
                                    (dst.Width / 2) + offset, dst.Height / 4);
                                break;
                            case 5:
                                pcb.AddTemplate(page,
                                    scale, 0f, 0f, scale,
                                    (dst.Width / 2) + offset, 0);
                                break;

                            // LEFT SIDE
                            case 6:
                                pcb.AddTemplate(page,
                                    -scale, 0f, 0f, -scale,
                                    (dst.Width / 2) - offset, dst.Height / 4);
                                break;
                            case 7:
                                pcb.AddTemplate(page,
                                    -scale, 0f, 0f, -scale,
                                    (dst.Width / 2) - offset, dst.Height * 2 / 4);
                                break;
                            case 0:
                                pcb.AddTemplate(page,
                                    -scale, 0f, 0f, -scale,
                                    (dst.Width / 2) - offset, dst.Height * 3 / 4);
                                break;
                        }

                    }
                    else
                    {
                        switch (i % 8)
                        {
                            // LEFT SIDE
                            case 1:
                                // Counterclockwise 90 degrees
                                pcb.AddTemplate(page,
                                    0f, scale, -scale, 0f,
                                    (dst.Width / 2) - offset, dst.Height * 3 / 4);
                                break;

                            // RIGHT SIDE
                            case 2:
                                // Clockwise 90 degrees
                                pcb.AddTemplate(page,
                                    0f, -scale, scale, 0f,
                                    (dst.Width / 2) + offset, dst.Height);
                                break;
                            case 3:
                                // Clockwise 90 degrees
                                pcb.AddTemplate(page,
                                    0f, -scale, scale, 0f,
                                    (dst.Width / 2) + offset, dst.Height * 3 / 4);
                                break;
                            case 4:
                                // Clockwise 90 degrees
                                pcb.AddTemplate(page,
                                    0f, -scale, scale, 0f,
                                    (dst.Width / 2) + offset, dst.Height * 2 / 4);
                                break;
                            case 5:
                                // Clockwise 90 degrees
                                pcb.AddTemplate(page,
                                    0f, -scale, scale, 0f,
                                    (dst.Width / 2) + offset, dst.Height / 4);
                                break;

                            // LEFT SIDE
                            case 6:
                                // Counterclockwise 90 degrees
                                pcb.AddTemplate(page,
                                    0f, scale, -scale, 0f,
                                    (dst.Width / 2) - offset, 0);
                                break;
                            case 7:
                                // Counterclockwise 90 degrees
                                pcb.AddTemplate(page,
                                    0f, scale, -scale, 0f,
                                    (dst.Width / 2) - offset, dst.Height / 4);
                                break;
                            case 0:
                                // Counterclockwise 90 degrees
                                pcb.AddTemplate(page,
                                    0f, scale, -scale, 0f,
                                    (dst.Width / 2) - offset, dst.Height * 2 / 4);
                                break;
                        }

                    }


                    if ((i % 8 == 0)||(i == pr.NumberOfPages))
                    {
                        // Draw guide line for folding.
                        pcb.SetLineWidth(0.01f);

                        if (checkBox1.Checked)
                        {
                            pcb.MoveTo(0f, 0f);
                            pcb.LineTo(dst.Width, 0f);
                            pcb.LineTo(dst.Width, dst.Height);
                            pcb.LineTo(0f, dst.Height);
                            pcb.LineTo(0f, 0f);
                            pcb.Stroke();
                        }

                        pcb.MoveTo(0f, dst.Height * 3 / 4f);
                        pcb.LineTo(dst.Width, dst.Height * 3 / 4);
                        pcb.Stroke();

                        pcb.MoveTo(0f, dst.Height * 2 / 4f);
                        pcb.LineTo(dst.Width, dst.Height * 2 / 4);
                        pcb.Stroke();

                        pcb.MoveTo(0f, dst.Height * 1 / 4f);
                        pcb.LineTo(dst.Width, dst.Height * 1 / 4);
                        pcb.Stroke();

                        pcb.MoveTo(dst.Width / 2, 0f);
                        pcb.LineTo(dst.Width / 2, dst.Height * 1 / 4);
                        pcb.Stroke();

                        pcb.MoveTo(dst.Width / 2, dst.Height * 3 / 4);
                        pcb.LineTo(dst.Width / 2, dst.Height);
                        pcb.Stroke();

                        pcb.SetLineDash(3f, 3f);
                        pcb.MoveTo(dst.Width / 2, dst.Height * 1 / 4);
                        pcb.LineTo(dst.Width / 2, dst.Height * 3 / 4);
                        pcb.Stroke();

                        pcb.SetLineDash(0);
                    }

                }

                doc.Close();
                pw.Close();
                pr.Close();

            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
            }

        }
        /*
                private void CheckPdfProtection(string filePath)
                {
                    try
                    {
                        PdfReader reader = new PdfReader(filePath);
                        if (!reader.IsEncrypted()) return;
                        if (!PdfEncryptor.IsPrintingAllowed((int)reader.Permissions))
                            throw new InvalidOperationException("the selected file is print protected and cannot be imported");
                        if (!PdfEncryptor.IsModifyContentsAllowed((int)reader.Permissions))
                            throw new InvalidOperationException("the selected file is write protected and cannot be imported");
                    }
                    //catch (BadPasswordException) { throw new InvalidOperationException("the selected file is password protected and cannot be imported"); }
                    catch (BadPdfFormatException) { throw new InvalidDataException("the selected file is having invalid format and cannot be imported"); }
                }
        */
    }
}


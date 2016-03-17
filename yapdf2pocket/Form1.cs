/*
Copyright (c) 2015-2016 acknpop
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

                var dst = PageSize.GetRectangle(comboBox1.SelectedItem.ToString());
                var doc = new Document(dst);
                var pw = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                doc.Open();

                for (var page = 1; page <= pr.NumberOfPages; page++)
                {
                    if (page % 8 == 1) doc.NewPage();

                    var pcb = pw.DirectContent;
                    var src = pr.GetPageSize(page);
                    int rot = pr.GetPageRotation(page);

                    bool isLandscape = (src.Width > src.Height) ? true : false;
                    float scale;
                    float offset;
                    if (isLandscape)
                    {
                        scale = dst.Height / (src.Height * 4);
                        offset = (dst.Width / 2 - (scale * src.Width)) / 2;
                    }
                    else
                    {
                        scale = dst.Height / (src.Width * 4);
                        offset = (dst.Width / 2 - (scale * src.Height)) / 2;
                    }

                    // For scaling and rotation
                    var transRotate = new iTextSharp.awt.geom.AffineTransform();
                    transRotate.SetToIdentity();
                    transRotate.Scale(scale, scale);

                    // For position
                    var transAdjust = new iTextSharp.awt.geom.AffineTransform();
                    transAdjust.SetToIdentity();

                    // Affine translation of PocketMod style.
                    if (isLandscape || (rot == 90 || rot == 270))
                    {
                        var px = dst.Width / 2;
                        var py = dst.Height / 4;
                        var rads = 180 * Math.PI / 180;

                        switch (page % 8)
                        {
                            // LEFT SIDE
                            // Rotate 180 degrees
                            case 1:
                                transAdjust.Translate(px - offset, py * 4);
                                transRotate.Rotate(-rads);
                                break;
                            case 0: // as 8
                                transAdjust.Translate(px - offset, py * 3);
                                transRotate.Rotate(-rads);
                                break;
                            case 7:
                                transAdjust.Translate(px - offset, py * 2);
                                transRotate.Rotate(-rads);
                                break;
                            case 6:
                                transAdjust.Translate(px - offset, py);
                                transRotate.Rotate(-rads);
                                break;

                            // RIGHT SIDE
                            // no rotation
                            case 2:
                                transAdjust.Translate(px + offset, py * 3);
                                break;
                            case 3:
                                transAdjust.Translate(px + offset, py * 2);
                                break;
                            case 4:
                                transAdjust.Translate(px + offset, py);
                                break;
                            case 5:
                                transAdjust.Translate(px + offset, 0);
                                break;
                        }

                    }
                    else
                    {
                        var px = dst.Width / 2;
                        var py = dst.Height / 4;
                        var rads = 90 * Math.PI / 180;

                        switch (page % 8)
                        {
                            // LEFT SIDE
                            // Rotate counterclockwise 90 degrees
                            case 1:
                                transAdjust.Translate(px - offset, py * 3);
                                transRotate.Rotate(rads);
                                break;
                            case 0: // as 8
                                transAdjust.Translate(px - offset, py * 2);
                                transRotate.Rotate(rads);
                                break;
                            case 7:
                                transAdjust.Translate(px - offset, py);
                                transRotate.Rotate(rads);
                                break;
                            case 6:
                                transAdjust.Translate(px - offset, 0);
                                transRotate.Rotate(rads);
                                break;

                            // RIGHT SIDE
                            // Rotate clockwise 90 degrees
                            case 2:
                                transAdjust.Translate(px + offset, py * 4);
                                transRotate.Rotate(-rads);
                                break;
                            case 3:
                                transAdjust.Translate(px + offset, py * 3);
                                transRotate.Rotate(-rads);
                                break;
                            case 4:
                                transAdjust.Translate(px + offset, py * 2);
                                transRotate.Rotate(-rads);
                                break;
                            case 5:
                                transAdjust.Translate(px + offset, py);
                                transRotate.Rotate(-rads);
                                break;
                        }

                    }
                    var finalTrans = new iTextSharp.awt.geom.AffineTransform();
                    finalTrans.SetToIdentity();
                    finalTrans.Concatenate(transAdjust);
                    finalTrans.Concatenate(transRotate);

                    var importedPage = pw.GetImportedPage(pr, page);
                    pcb.AddTemplate(importedPage, finalTrans);

                    if ((page % 8 == 0)||(page == pr.NumberOfPages))
                    {
                        // Draw guide line for folding.
                        pcb.SetLineWidth(0.01f);

                        // Outside frame
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


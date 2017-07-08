using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using GemBox.Document;

namespace WindowsFormsRichTextEditor
{
    public partial class MainForm : Form
    {
        private readonly List<byte[]> sampleFiles = new List<byte[]>()
        {
            Resources.Character_Formatting,
            Resources.Paragraph_Formatting,
            Resources.Lists,
            Resources.Style_Resolution,
            Resources.Simple_Table,
        };

        public MainForm()
        {
            InitializeComponent();

            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        #region Event Handlers

        private void menuItemOpen_Click(object sender, EventArgs e)
        {
            var dialog = new OpenFileDialog()
            {
                AddExtension = true,
                Filter =
                    "All Documents (*.docx;*.docm;*.doc;*.dotx;*.dotm;*.dot;*.htm;*.html;*.rtf;*.txt)|*.docx;*.docm;*.dotx;*.dotm;*.doc;*.dot;*.htm;*.html;*.rtf;*.txt|" +
                    "Word Documents (*.docx)|*.docx|" +
                    "Word Macro-Enabled Documents (*.docm)|*.docm|" +
                    "Word 97-2003 Documents (*.doc)|*.doc|" +
                    "Word Templates (*.dotx)|*.dotx|" +
                    "Word Macro-Enabled Templates (*.dotm)|*.dotm|" +
                    "Word 97-2003 Templates (*.dot)|*.dot|" +
                    "Web Pages (*.htm;*.html)|*.htm;*.html|" +
                    "Rich Text Format (*.rtf)|*.rtf|" +
                    "Text Files (*.txt)|*.txt"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                using (var stream = new MemoryStream())
                {
                    // Convert input file to RTF stream.
                    DocumentModel.Load(dialog.FileName).Save(stream, SaveOptions.RtfDefault);

                    stream.Position = 0;

                    // Load RTF stream into RichTextBox.
                    this.richTextBox.LoadFile(stream, RichTextBoxStreamType.RichText);
                }
        }

        private void menuItemOpenSample_Click(object sender, EventArgs e)
        {
            var menuItem = sender as ToolStripMenuItem;
            if (menuItem != null)
            {
                int sampleFileIndex = 0;
                string tag = menuItem.Tag as string;
                if(!string.IsNullOrEmpty(tag))
                {
                    switch (tag)
                    {
                        case "1":
                            sampleFileIndex = 0;
                            break;
                        case "2":
                            sampleFileIndex = 1;
                            break;
                        case "3":
                            sampleFileIndex = 2;
                            break;
                        case "4":
                            sampleFileIndex = 3;
                            break;
                        case "5":
                            sampleFileIndex = 4;
                            break;
                        default:
                            throw new NotImplementedException();
                    }
                }

                DocumentModel doc;
                using (MemoryStream stream = new MemoryStream(sampleFiles[sampleFileIndex]))
                    doc = DocumentModel.Load(stream, LoadOptions.DocxDefault);

                using (MemoryStream stream = new MemoryStream())
                {
                    doc.Save(stream, DocxSaveOptions.RtfDefault);
                    stream.Seek(0, SeekOrigin.Begin);

                    // Load RTF stream into RichTextBox.
                    this.richTextBox.LoadFile(stream, RichTextBoxStreamType.RichText);
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            var dialog = new SaveFileDialog()
            {
                AddExtension = true,
                Filter =
                    "Word Document (*.docx)|*.docx|" +
                    "Word Macro-Enabled Document (*.docm)|*.docm|" +
                    "Word Template (*.dotx)|*.dotx|" +
                    "Word Macro-Enabled Template (*.dotm)|*.dotm|" +
                    "PDF (*.pdf)|*.pdf|" +
                    "XPS Document (*.xps)|*.xps|" +
                    "Web Page (*.htm;*.html)|*.htm;*.html|" +
                    "Single File Web Page (*.mht;*.mhtml)|*.mht;*.mhtml|" +
                    "Rich Text Format (*.rtf)|*.rtf|" +
                    "Plain Text (*.txt)|*.txt|" +
                    "Image (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.wdp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.wdp"
            };

            if (dialog.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                using (var stream = new MemoryStream())
                {
                    // Save RichTextBox content to RTF stream.
                    this.richTextBox.SaveFile(stream, RichTextBoxStreamType.RichText);

                    stream.Position = 0;

                    // Convert RTF stream to output format.
                    DocumentModel.Load(stream, LoadOptions.RtfDefault).Save(dialog.FileName);
                    Process.Start(dialog.FileName);
                }
        }

        private void btnGemBoxCut_Click(object sender, EventArgs e)
        {
            this.DoGemBoxCopy();

            // Clear selection.
            this.richTextBox.SelectedRtf = string.Empty;
        }

        private void btnGemBoxCopy_Click(object sender, EventArgs e)
        {
            this.DoGemBoxCopy();
        }

        private void btnGemBoxPastePrepend_Click(object sender, EventArgs e)
        {
            this.DoGemBoxPaste(true);
        }

        private void btnGemBoxPasteAppend_Click(object sender, EventArgs e)
        {
            this.DoGemBoxPaste(false);
        }

        private void btnUndo_Click(object sender, EventArgs e)
        {
            if (this.richTextBox.CanUndo)
                this.richTextBox.Undo();
        }

        private void btnRedo_Click(object sender, EventArgs e)
        {
            if (this.richTextBox.CanRedo)
                this.richTextBox.Redo();
        }

        private void btnCut_Click(object sender, EventArgs e)
        {
            this.richTextBox.Cut();
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            this.richTextBox.Copy();
        }

        private void btnPaste_Click(object sender, EventArgs e)
        {
            this.richTextBox.Paste();
        }

        private void btnBold_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionFont = new Font(this.richTextBox.SelectionFont.FontFamily, this.richTextBox.SelectionFont.Size, this.ToggleFontStyle(this.richTextBox.SelectionFont.Style, FontStyle.Bold));
        }

        private void btnItalic_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionFont = new Font(this.richTextBox.SelectionFont.FontFamily, this.richTextBox.SelectionFont.Size, this.ToggleFontStyle(this.richTextBox.SelectionFont.Style, FontStyle.Italic));
        }

        private void btnUnderline_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionFont = new Font(this.richTextBox.SelectionFont.FontFamily, this.richTextBox.SelectionFont.Size, this.ToggleFontStyle(this.richTextBox.SelectionFont.Style, FontStyle.Underline));
        }

        private void btnIncreaseFontSize_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionFont = new Font(this.richTextBox.SelectionFont.FontFamily, this.richTextBox.SelectionFont.Size + 1, this.richTextBox.SelectionFont.Style);
        }

        private void btnDecreaseFontSize_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionFont = new Font(this.richTextBox.SelectionFont.FontFamily, this.richTextBox.SelectionFont.Size - 1, this.richTextBox.SelectionFont.Style);
        }

        private void btnToggleBullets_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionBullet = !this.richTextBox.SelectionBullet;
        }

        private void btnDecreaseIndentation_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionIndent -= 10;
        }

        private void btnIncreaseIndentation_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionIndent += 10;
        }

        private void btnAlignLeft_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Left;
        }

        private void btnAlignCenter_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Center;
        }

        private void btnAlignRight_Click(object sender, EventArgs e)
        {
            this.richTextBox.SelectionAlignment = System.Windows.Forms.HorizontalAlignment.Right;
        }

        #endregion

        private void DoGemBoxCopy()
        {
            using (var stream = new MemoryStream())
            {
                // Save RichTextBox selection to RTF stream.
                var writer = new StreamWriter(stream);
                writer.Write(this.richTextBox.SelectedRtf);
                writer.Flush();

                stream.Position = 0;

                // Save RTF stream to clipboard.
                DocumentModel.Load(stream, LoadOptions.RtfDefault).Content.SaveToClipboard();
            }
        }

        private void DoGemBoxPaste(bool prepend)
        {
            using (var stream = new MemoryStream())
            {
                // Save RichTextBox content to RTF stream.
                var writer = new StreamWriter(stream);
                writer.Write(this.richTextBox.SelectedRtf);
                writer.Flush();

                stream.Position = 0;

                // Load document from RTF stream and prepend or append clipboard content to it.
                var document = DocumentModel.Load(stream, LoadOptions.RtfDefault);
                var content = document.Content;
                (prepend ? content.Start : content.End).LoadFromClipboard();

                stream.Position = 0;

                // Save document to RTF stream.
                document.Save(stream, SaveOptions.RtfDefault);

                stream.Position = 0;

                // Load RTF stream into RichTextBox.
                var reader = new StreamReader(stream);
                this.richTextBox.SelectedRtf = reader.ReadToEnd();
            }
        }

        private FontStyle ToggleFontStyle(FontStyle item, FontStyle toggle)
        {
            return item ^ toggle;
        }

    }
}

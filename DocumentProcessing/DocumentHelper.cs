using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ExamProject
{
    public class DocumentHelper
    {
        private FileInfo fileInfo;
        public DocumentHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                fileInfo = new FileInfo(fileName);
            }
            else
            {
                MessageBox.Show("File not found", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal bool Process(Dictionary<string, string> items)
        {
            if (items.ContainsValue(string.Empty))
            {
                MessageBox.Show("Пожалуйста, введите все необходимые данные.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = fileInfo.FullName;
                Object missing = Type.Missing;
                app.Documents.Open(file);
                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing,
                        Replace: replace);

                }

                Object newFileName = Path.Combine(fileInfo.DirectoryName, DateTime.Now.ToString("yyyyMMdd HHmmss") + fileInfo.Name);
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();
                app.Quit();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }
            return false;
        }
    }
    
}

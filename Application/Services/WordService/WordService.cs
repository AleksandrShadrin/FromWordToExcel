using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace Application.Services.WordService
{
    public class WordService : IDisposable
    {
        private Word.Application _application;
        private Word.Document _document;
        private StringBuilder _stringBuilder = new();
        private WordService(Word.Application application, Word.Document document)
        {
            _application = application;
            _document = document;
        }
        public void Dispose()
        {
            _document.Close();
            Marshal.ReleaseComObject(_document);

            _application.Quit();
            Marshal.ReleaseComObject(_application);

            GC.Collect();
        }
        public IEnumerable<string> ReadLines()
        {
            int count = 1;
            string word;
            while (count <= _document.Words.Count)
            {
                word = _document.Words[count].Text;
                _stringBuilder.Append(word);
                count++;
                if (Char.TryParse(word, out var res))
                {
                    if ((int)res == 13)
                    {
                        var line = _stringBuilder.ToString();
                        _stringBuilder.Clear();
                        yield return line;
                    }
                }
            }
        }
        public static WordService OpenWordFile(string filePath)
        {
            Word.Application application = null;
            Word.Document document = null;
            try
            {
                application = new();
                document = application.Documents.Open(filePath);
                return new WordService(application, document);
            }
            catch (Exception)
            {
                if (document != null)
                {
                    document.Close();
                    Marshal.ReleaseComObject(document);
                }
                if (application != null)
                {
                    application.Quit();
                    Marshal.ReleaseComObject(application);
                }
                throw;
            }
        }
    }
}

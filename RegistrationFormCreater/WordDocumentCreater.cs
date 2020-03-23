using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;

namespace RegistrationFormCreater
{
    public sealed class FormOption
    { 
        public byte[] TemplateContent { get; set; }
    }

    public class WordDocumentCreater
    {
        private readonly char _bookmarkSuffix = '_';
        private readonly FormOption _option;

        public WordDocumentCreater(FormOption option)
        {
            this._option = option;
        }

        public byte[] GetFile(Form item)
        {
            var _templateBuffer = this._option.TemplateContent;
            var _outBuffer = new byte[] { };

            var _keyValues = this.ToKeyValues(item);

            using (MemoryStream _stream = new MemoryStream())
            {
                _stream.Write(_templateBuffer, 0, _templateBuffer.Length);

                var _document = WordprocessingDocument.Open(_stream, true);
                var _body = _document.MainDocumentPart.Document.Body;

                var _bookmarkName = string.Empty;
                var _key = string.Empty;
                var _value = string.Empty;

                foreach (BookmarkStart _bookmarkStart in _body.Descendants<BookmarkStart>())
                {
                    _bookmarkName = _bookmarkStart.Name;

                    //緊處理特定標籤結尾的標籤
                    if (_bookmarkName.EndsWith(_bookmarkSuffix) == false)
                        continue;

                    _key = _bookmarkName.TrimEnd(new char[] { _bookmarkSuffix });
                    _value = string.Empty;

                    if (_keyValues.ContainsKey(_key) == true)
                        _value = _keyValues[_key];

                    var _element = _bookmarkStart.Parent;

                    //字型設定
                    var _runProperty = new RunProperties();
                    var _runFont = new RunFonts()
                    {
                        Ascii = "Times New Roman",
                        EastAsia = "標楷體"
                    };
                    _runProperty.AppendChild(_runFont);
                    var _run = _element.AppendChild(new Run());

                    //文字加入字型
                    _run.PrependChild(_runProperty);
                    _run.AppendChild(new Text(_value));
                }

                _document.Close();

                _outBuffer = _stream.ToArray();
            }

            return _outBuffer;
        }

        private IDictionary<string, string> ToKeyValues(Form item)
        {
            Dictionary<string, string> _keyValues;

            _keyValues = new Dictionary<string, string>();
            _keyValues.Add("TeamName", item.TeamName);
            _keyValues.Add("Level", item.Level);
            _keyValues.Add("LeaderName", item.LeaderName);
            _keyValues.Add("LeaderPhone", item.LeaderPhone);
            _keyValues.Add("CoachName", item.CoachName);
            _keyValues.Add("CoachPhone", item.CoachPhone);

            var _index = 0;

            foreach (var _player in item.Players)
            {
                _keyValues.Add($"Number{_index}", _player.Number);
                _keyValues.Add($"Name{_index}", _player.Name);
                _keyValues.Add($"CaptionMark{_index}", _player.CaptionMark);
                _keyValues.Add($"GoalKeeperMark{_index}", _player.GoalKeeperMark);
                _keyValues.Add($"Note{_index}", _player.Note);

                _index = _index + 1;
            }

            return _keyValues;
        }
    }

}

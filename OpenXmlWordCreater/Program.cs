using RegistrationFormCreater;
using System;
using System.IO;

namespace OpenXmlWordCreater
{
    class Program
    {
        //參考資料
        //https://zh.wikipedia.org/wiki/%E4%B8%AD%E8%8F%AF%E5%8F%B0%E5%8C%97%E4%BA%94%E4%BA%BA%E5%88%B6%E8%B6%B3%E7%90%83%E4%BB%A3%E8%A1%A8%E9%9A%8A#%E6%9C%80%E8%BF%91%E5%85%A5%E9%81%B8%E7%90%83%E5%93%A1%E5%90%8D%E5%96%AE

        const string _templatePath = "../../../../files/template.docx";
        const string _outPath = "../../../../files/out.docx";

        static void Main(string[] args)
        {
            Form _form = new Form();
            _form.TeamName = string.Format("中華隊 {0:yyyy/MM/dd HH:mm:ss}", DateTime.UtcNow);
            _form.Level = "2012亞足聯五人制足球賽";
            _form.LeaderName = "蔡O峰";
            _form.LeaderPhone = "0900-000-000";
            _form.CoachName = "強O在";
            _form.CoachPhone = "0998-000-022";
            _form.Players = new PlayerItem[] {
                new PlayerItem(){ Number = "1", Name = "潘O傑", IsCaption = false, IsGoalKeeper = true   , Note = "Note01" },
                new PlayerItem(){ Number = "12", Name = "施O安", IsCaption = false, IsGoalKeeper = true  , Note = "Note02" },
                new PlayerItem(){ Number = "2", Name = "謝O軒", IsCaption = false, IsGoalKeeper = false  , Note = "Note03" },
                new PlayerItem(){ Number = "3", Name = "翁O賓", IsCaption = false, IsGoalKeeper = false  , Note = "Note04" },
                new PlayerItem(){ Number = "4", Name = "陳O瀚", IsCaption = false, IsGoalKeeper = false  , Note = "Note05" },
                new PlayerItem(){ Number = "6", Name = "黃O宗", IsCaption = true, IsGoalKeeper = false   , Note = "Note06" },
                new PlayerItem(){ Number = "8", Name = "羅O恩", IsCaption = false, IsGoalKeeper = false  , Note = "Note07" },
                new PlayerItem(){ Number = "9", Name = "羅O安", IsCaption = false, IsGoalKeeper = false  , Note = "Note08" },
                new PlayerItem(){ Number = "11", Name = "陳O豪", IsCaption = false, IsGoalKeeper = false , Note = "Note09" },
                new PlayerItem(){ Number = "14", Name = "劉O超", IsCaption = false, IsGoalKeeper = false , Note = "Note10" },
                new PlayerItem(){ Number = "16", Name = "楊O勛", IsCaption = false, IsGoalKeeper = false , Note = "Note11" },
                new PlayerItem(){ Number = "17", Name = "陳O楠", IsCaption = false, IsGoalKeeper = false , Note = "Note12" },
                new PlayerItem(){ Number = "19", Name = "齊O棋", IsCaption = false, IsGoalKeeper = false , Note = "Note13" }
            };

            FormOption _option;
            _option = new FormOption();
            _option.TemplateContent = File.ReadAllBytes(_templatePath);

            WordDocumentCreater _create = new WordDocumentCreater(_option);
            var _output = _create.GetFile(_form);

            File.WriteAllBytes(_outPath, _output);
        }
    }
}

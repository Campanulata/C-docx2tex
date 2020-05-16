using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.Formula.Functions;
using System.Text.RegularExpressions;
using SqlSugar;

namespace cSharpLatex
{
    class Choice
    {
        public int Id { get; set; }
        public string Year { get; set; }
        public string City { get; set; }
        public string Qst { get; set; }
        public int WithChart { get; set; }
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string Answer { get; set; }
        public Choice()
        {
            Answer = "C";
        }
    }
    class UnChoice
    {
        public int Id { get; set; }
        public string Year { get; set; }
        public string City { get; set; }
        public string Qst { get; set; }
        public int WithChart { get; set; }
        public string Sub1 { get; set; }
        public string Sub2 { get; set; }
        public string Sub3 { get; set; }
        public string Sub4 { get; set; }
        public string Sub5 { get; set; }
        public string Sub6 { get; set; }
        public string Sub7 { get; set; }
    }
    class Paper
    {
        public string Year { get; set; }
        public string City { get; set; }
        public List<int> WithChart { get; set; }
        public List<int> QuestionNo { get; set; }
        public int CountAll { get; set; }
        public List<string> QuestionNoStr { get; set; }
        public XWPFDocument Docx { get; set; }
        public string PreChoice { get; set; }
        public IList<XWPFParagraph> Para { get; set; }
        public bool IsChoice { get; set; }
        Dictionary<string, string> cityCode = new Dictionary<string, string>{
            { "全国I","1" },
            { "全国Ⅱ","2"},
            { "全国Ⅲ","3"},
            { "北京","4"},
            { "江苏","5"},
            { "浙江","6"},
            { "海南","7"},
            { "天津","8"}
        };
        public void ShowInfo(Paper paper)
        {
            Console.WriteLine(paper.Year + paper.City);
            Console.WriteLine(paper.QuestionNo);
            Console.WriteLine(QuestionNoStr);
        }
        /// <summary>
        /// 初始化试卷类，传入一个XWPFDocument类型
        /// </summary>
        /// <param name="doc">NPOI的XWPFDocument类型</param>
        public Paper(XWPFDocument doc)
        {
            Docx = doc;
            Para = Docx.Paragraphs;
            Year = Para[0].Text.Substring(0, 4);
            for (int i = doc.Paragraphs.Count - 1; i > 0; i--)
            {
                if (doc.Paragraphs[i].Text == "")
                {
                    continue;
                }
                string mm = doc.Paragraphs[i].Text.Substring(0, 2);
                if (Regex.IsMatch(mm, @"^[-+]?\d+(\.\d+)?$"))
                {
                    CountAll=int.Parse(mm);
                    break;
                }
            }

        }
        /// <summary>
        /// 增加试卷识别的容错率，手动设定第一道题目上一行字符串的前两个字符,选择题数量和总数
        /// </summary>
        /// <param name="a">常见类型：一、</param>
        /// <param name="b">选择题数</param>
        /// <param name="filePath">包含文件名的绝对路径，判断city</param>
        public void SetPreChoiceAndCount(string a,string filePath)
        {
            for (int i = 0; i < 10; i++)
            {
                QuestionNoStr.Add(i + ".");
                QuestionNo.Add(0);
                WithChart.Add(0);
            }
            for (int i = 10; i < CountAll + 1; i++)
            {
                QuestionNoStr.Add(i.ToString());
                QuestionNo.Add(0);
                WithChart.Add(0);
            }
            QuestionNoStr[0] = a;
            foreach (var k in cityCode.Keys)
            {
                if (filePath.Contains(k))
                {
                    City = cityCode[k];
                    break;
                }
            }
        }  
        public void GetIsChoice(int i)
        {
            string concatAll = "";
            if (i==CountAll)
            {
                IsChoice = false;
            }
            else
            {
                for (int j = QuestionNo[i] + 1; j < QuestionNo[i + 1]; j++)//本题除了题干都caoncat
                {
                    if (Para[j].Text == "")
                    {
                        continue;
                    }
                    concatAll = concatAll + Para[j].Text;
                }
                if (concatAll.Contains("A") && concatAll.Contains("B") && concatAll.Contains("C") && concatAll.Contains("D"))
                {
                    IsChoice = true;
                }
            }


        }
        /// <summary>
        /// 让第i题的信息赋给choice/unChoice对象，返回这个对象
        /// </summary>
        /// <param name="i">题号</param>
        public Choice ChoiceToObj(int i)
        {
            string id = i.ToString();
            if (i < 10)
            {
                id = "0" + id;
            }
            int[] abcd = { 0, 0, 0, 0 };
            for (int j = QuestionNo[i]; j < QuestionNo[i + 1]; j++)
            {
                if (Para[j].Text == "" || Para[j+1].Text == "")
                {
                    continue;
                }
                if (Para[j].Text.Substring(0, 2) == "A." && Para[j + 3].Text.Substring(0, 2) == "D.")
                {
                    for (int f = 0; f < 4; f++)
                    {
                        abcd[f] = j + f;
                    }
                    break;
                }
                else if (Para[j].Text.Substring(0, 2) == "A." && Para[j + 1].Text.Substring(0, 2) == "C.")
                {
                    abcd[0] = j; abcd[1] = j; abcd[2] = j + 1; abcd[3] = j + 1;
                    break;
                }
            }
            Choice choice = new Choice();
            choice.Id = int.Parse(Year + City + id); choice.Year = Year; choice.City = City; choice.WithChart = WithChart[i];
            choice.Qst = Para[QuestionNo[i]].Text;
            if (choice.WithChart == 1)
            {
                choice.Qst = choice.Qst + @"\begin{center}\includegraphics[width=4cm]{./pic/" + choice.Id + @"0.PNG}\end{center}";
            }
            choice.A = Para[abcd[0]].Text; choice.B = Para[abcd[1]].Text; choice.C = Para[abcd[2]].Text; choice.D = Para[abcd[3]].Text;
            return choice;
        }
        public UnChoice UnChoiceToObj(int i)
        {
            List<int> SubList = new List<int> { 0, 0, 0, 0, 0, 0, 0, 0 };SubList[0] = QuestionNo[i];
            string id = i.ToString();
            if (i < 10)
            {
                id = "0" + id;
            }
            UnChoice unChoice = new UnChoice();
            unChoice.Id = int.Parse(Year + City + id); unChoice.Year = Year; unChoice.City = City; unChoice.WithChart = WithChart[i];

            for (int j = QuestionNo[i] + 1; j < QuestionNo[i + 1]; j++)
            {
                if (Para[j].Text == "")
                {
                    continue;
                }
                string mmm = Para[j].Text.Substring(0, 3);
                switch (mmm)
                {
                    case "（1）":
                        unChoice.Sub1 = Para[j].Text;SubList[1] = j;
                        break;
                    case "（2）":
                        unChoice.Sub2 = Para[j].Text; SubList[2] = j;
                        break;
                    case "（3）":
                        unChoice.Sub3 = Para[j].Text; SubList[3] = j;
                        break;
                    case "（4）":
                        unChoice.Sub4 = Para[j].Text; SubList[4] = j;
                        break;
                    case "（5）":
                        unChoice.Sub5 = Para[j].Text; SubList[5] = j;
                        break;
                    case "（6）":
                        unChoice.Sub6 = Para[j].Text; SubList[6] = j;
                        break;
                    case "（7）":
                        unChoice.Sub7 = Para[j].Text; SubList[7] = j;
                        break;
                    default:
                        break;
                }
            }
            for (int k = SubList[0]; k < SubList[1]; k++)
            {
                unChoice.Qst = unChoice.Qst + Para[k].Text;
            }
            if (unChoice.WithChart == 1)
            {
                unChoice.Qst = unChoice.Qst + @"\begin{center}\includegraphics[width=4cm]{./pic/" + unChoice.Id + @"0.PNG}\end{center}";
            }
            return unChoice;
        }
 
        public void QstToDB()
        {
            SqlSugarClient db = new SqlSugarClient(
                new ConnectionConfig()
                {
                    ConnectionString = @"DataSource=E:\微云\C#\cSharpLatex\Question.db",
                    DbType = DbType.Sqlite,//设置数据库类型
                    IsAutoCloseConnection = true,//自动释放数据务，如果存在事务，在事务结束后释放
                    InitKeyType = InitKeyType.Attribute //从实体特性中读取主键自增列信息
                });
            
            
            for (int i = 1; i < CountAll +1; i++)
            {
                GetIsChoice(i);
                if (i>21)
                {
                    IsChoice = false;
                }
                if (IsChoice)
                {
                    //Choice choice = new Choice();
                    //choice = ChoiceToObj(i);
                    //db.Insertable(choice).ExecuteCommand();
                    db.Insertable(ChoiceToObj(i)).ExecuteCommand();

                    IsChoice = false;
                }
                else
                {
                    db.Insertable(UnChoiceToObj(i)).ExecuteCommand();
                    IsChoice = false;

                }
            }
        }
        /// <summary>
        /// 得到每道题的行号
        /// </summary>
        public void GetQuestionNo()
        {
            for (int i = 0; i < CountAll + 1; i++)//从0到countAll循环一边,每次循环可以给第i题赋值
            {
                int preNo = 0;
                if (i>0)
                {
                    preNo = QuestionNo[i - 1];
                }

                for (int j = preNo; j < Para.Count+1; j++)//按para的行循环一边
                {
                    if (Para[j].Text == "")
                    {
                        WithChart[i] = 1;
                        continue;
                    }
                    if (QuestionNo[i]>0  )
                    {
                        break;
                    }
                    else if (Para[j].Text.Substring(0, 2) == QuestionNoStr[i])
                    {
                        QuestionNo[i] = j;
                    }

                }

            }
            QuestionNo.Add(QuestionNo[CountAll]);
        }
    }
}

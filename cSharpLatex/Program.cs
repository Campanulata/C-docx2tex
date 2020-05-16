using System;
using CSharpMath;
using NPOI;
using NPOI.XWPF.Model;
using NPOI.XWPF.Extractor;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using SqlSugar;
using System.Text.RegularExpressions;

namespace cSharpLatex
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"E:\微云\C#\cSharpLatex\cSharpLatex\2019年全国I卷理科综合高考真题.docx";
            FileStream filestream = File.OpenRead(filePath);
            XWPFDocument doc = new XWPFDocument(filestream);

            Paper paper = new Paper(doc)
            {
                QuestionNo = new List<int>(),QuestionNoStr = new List<string>(),WithChart=new List<int>()
            };
            paper.SetPreChoiceAndCount("一、", filePath);
            paper.GetQuestionNo();

            paper.QstToDB();
            //对象=>DB
            Console.WriteLine("1");

        }
    }
}

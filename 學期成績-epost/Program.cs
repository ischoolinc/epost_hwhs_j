﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.Presentation.Controls;

namespace hwhs_epost_semester
{
    // 2019/11/12 穎驊註解 本專案為弘文於本學期提出來的需求 希望 能將 定期評量通知單(測試版) 輸出CSV檔支援 epost 列印
    // 製作方向 參照原本報表邏輯、新固定排名資料抓法、缺曠通知單的抓法， 以缺曠通知單 流程為基底修改而成
    // 2020/04/09 學期成績epsot 以之前完成的 定期評量通知單epsot 改寫，因為功能變數為通用
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            string URL學生學期成績通知單 = "ischool/高中系統/共用/學務/學生/報表/學期成績通知單_2019_弘文epost";            

            string toolName = "學期成績通知單(弘文epost)";

            FISCA.Features.Register(URL學生學期成績通知單, arg =>
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    new Report("student").Print(K12.Presentation.NLDPanels.Student.SelectedSource);
                }
                else
                {
                    MsgBox.Show("產生學生報表,請選擇學生!!");
                }
            });

            RibbonBarItem StudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            StudentReports["報表"]["成績相關報表"][toolName].Enable = Permissions.學生學期成績通知單權限;
            StudentReports["報表"]["成績相關報表"][toolName].Click += delegate
            {
                Features.Invoke(URL學生學期成績通知單);
            };


            //學生選擇
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                {
                    StudentReports["報表"]["成績相關報表"][toolName].Enable = false;
                }
                else
                {
                    StudentReports["報表"]["成績相關報表"][toolName].Enable = Permissions.學生學期成績通知單權限;
                }
            };


            Catalog ribbon = RoleAclSource.Instance["學生"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.學生學期成績通知單, toolName));

        }
    }
}

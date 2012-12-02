<%
'=================================================
'函数名：NewGetSlidePicArticle
'作  用：以幻灯片效果显示图片文章，此文章插入到PowerEasy.Article.asp 文件中
'参  数：
'0        iChannelID ---- 频道ID
'1        arrClassID ---- 栏目ID数组，0为所有栏目
'2        IncludeChild ---- 是否包含子栏目，仅当arrClassID为单个栏目ID时才有效，True----包含子栏目，False----不包含
'3        iSpecialID ---- 专题ID，0为所有文章（含非专题文章），如果为大于0，则只显示相应专题的文章
'4        ArticleNum ---- 最多显示多少篇文章
'5        IsHot ---- 是否是热门文章
'6        IsElite ---- 是否是推荐文章
'7        DateNum ---- 日期范围，如果大于0，则只显示最近几天内更新的文章
'8        OrderType ---- 排序方式，1--按文章ID降序，2--按文章ID升序，3--按更新时间降序，4--按更新时间升序，5--按点击数降序，6--按点击数升序，7--按评论数降序，8--按评论数升序
'9        ImgWidth ---- 图片宽度
'10       ImgHeight ---- 图片高度
'11       TitleLen ---- 文章标题字数限制，0为不显示，-1为显示完整标题
'12       iTimeOut ---- 效果变换间隔时间，以毫秒为单位
'13       effectID ---- 图片转换效果，0至22指定某一种特效，23表示随机效果
'=================================================
Public Function NewGetSlidePicArticle(iChannelID, arrClassID, IncludeChild, iSpecialID, ArticleNum, IsHot, IsElite, DateNum, OrderType, ImgWidth, ImgHeight, TitleLen, iTimeOut, effectID)
    Dim sqlPic, rsPic, i, strPic
    Dim DefaultPicUrl, strTitle

    ArticleNum = PE_CLng(ArticleNum)
    ImgWidth = PE_CLng(ImgWidth)
    ImgHeight = PE_CLng(ImgHeight)

    If ArticleNum <= 0 Or ArticleNum > 100 Then ArticleNum = 10
    If ImgWidth < 0 Or ImgWidth > 1000 Then ImgWidth = 150
    If ImgHeight < 0 Or ImgHeight > 1000 Then ImgHeight = 150
    If iTimeOut < 1000 Or iTimeOut > 100000 Then iTimeOut = 5000
    If effectID < 0 Or effectID > 23 Then effectID = 23

    FoundErr = False
    If (PE_Clng(iChannelID) <> 0 and Instr(iChannelID,",") = 0) and (PE_Clng(iChannelID)<>PrevChannelID Or ChannelID = 0) Then
        Call GetChannel(PE_Clng(iChannelID))
        PrevChannelID = iChannelID       
    End If  
    If FoundErr = True Then
        NewGetSlidePicArticle = ErrMsg
        Exit Function
    End If

    sqlPic = "select top " & ArticleNum & " A.ChannelID,A.ClassID,A.ArticleID,A.Title,A.UpdateTime,A.InfoPurview,A.InfoPoint,A.DefaultPicUrl"
    sqlPic = sqlPic & ",C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview"
    sqlPic = sqlPic & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, "", DateNum, OrderType, False, True)


    Dim ranNum
    Randomize
    ranNum = Int(900 * Rnd) + 100
    'strPic = "<script language=JavaScript>" & vbCrLf
    'strPic = strPic & "<!--" & vbCrLf
    'strPic = strPic & "var SlidePic_" & ranNum & " = new SlidePic_Article(""SlidePic_" & ranNum & """);" & vbCrLf
    'strPic = strPic & "SlidePic_" & ranNum & ".Width    = " & ImgWidth & ";" & vbCrLf
    'strPic = strPic & "SlidePic_" & ranNum & ".Height   = " & ImgHeight & ";" & vbCrLf
    'strPic = strPic & "SlidePic_" & ranNum & ".TimeOut  = " & iTimeOut & ";" & vbCrLf
    'strPic = strPic & "SlidePic_" & ranNum & ".Effect   = " & effectID & ";" & vbCrLf
    'strPic = strPic & "SlidePic_" & ranNum & ".TitleLen = " & TitleLen & ";" & vbCrLf
    PrevChannelID=0

    Set rsPic = Server.CreateObject("ADODB.Recordset")
    rsPic.Open sqlPic, Conn, 1, 1
    Do While Not rsPic.EOF
        'If iChannelID = 0 Then
            If rsPic("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsPic("ChannelID"))
                PrevChannelID = rsPic("ChannelID")
            End If
        'End If
        If Left(rsPic("DefaultPicUrl"), 1) <> "/" And InStr(rsPic("DefaultPicUrl"), "://") <= 0 Then
            DefaultPicUrl = ChannelUrl & "/" & UploadDir & "/" & rsPic("DefaultPicUrl")
        Else
            DefaultPicUrl = rsPic("DefaultPicUrl")
        End If
        If TitleLen = -1 Then
            strTitle = rsPic("Title")
        Else
            strTitle = GetSubStr(rsPic("Title"), TitleLen, ShowSuspensionPoints)
        End If
        
        'strPic = strPic & "var oSP = new objSP_Article();" & vbCrLf
        'strPic = strPic & "oSP.ImgUrl         = """ & DefaultPicUrl & """;" & vbCrLf
        'strPic = strPic & "oSP.LinkUrl        = """ & GetArticleUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("ArticleID"), rsPic("ClassPurview"), rsPic("InfoPurview"), rsPic("InfoPoint")) & """;" & vbCrLf
        'strPic = strPic & "oSP.Title         = """ & strTitle & """;" & vbCrLf
        'strPic = strPic & "SlidePic_" & ranNum & ".Add(oSP);" & vbCrLf

        '新的输出图片信息
        strPic = "<div id=""pic_news_info"">"& strTitle &"</div>" & vbCrLf
        strPic = strPic & "<div id=""pic_news_list"">" & vbCrLf
        strPic = strPic & "<a href=""http://www.usth.edu.cn"& GetArticleUrl(rsPic("ParentDir"), rsPic("ClassDir"), rsPic("UpdateTime"), rsPic("ArticleID"), rsPic("ClassPurview"), rsPic("InfoPurview"), rsPic("InfoPoint")) &""" target=""_blank""><img src="""& DefaultPicUrl &""" title="""& strTitle &""" alt="""& strTitle &""" /></a>" &vbCrLf
        strPic= strPic & "</div>" & vbCrLf
        
        rsPic.MoveNext
    Loop
    'strPic = strPic & "SlidePic_" & ranNum & ".Show();" & vbCrLf
    'strPic = strPic & "//-->" & vbCrLf
    'strPic = strPic & "</script>" & vbCrLf
    
    rsPic.Close
    Set rsPic = Nothing
    NewGetSlidePicArticle = strPic
End Function

'=================================================
'函数名：NewGetArticleList
'作  用：显示文章标题等信息
'参  数：
'0        iChannelID ---- 频道ID
'1        arrClassID ---- 栏目ID数组，0为所有栏目
'2        IncludeChild ---- 是否包含子栏目，仅当arrClassID为单个栏目ID时才有效，True----包含子栏目，False----不包含
'3        iSpecialID ---- 专题ID，0为所有文章（含非专题文章），如果为大于0，则只显示相应专题的文章
'4        UrlType ---- 链接地址类型，0为相对路径，1为带网址的绝对路径，不对外公开，4.03时为ShowAllArticle
'5        ArticleNum ---- 文章数，若大于0，则只查询前几篇文章
'6        IsHot ---- 是否是热门文章，True为只显示热门文章，False为显示所有文章
'7        IsElite ---- 是否是推荐文章，True为只显示推荐文章，False为显示所有文章
'8        Author ---- 作者姓名，如果不为空，则只显示指定作者的文章，用于作者文集
'9        DateNum ---- 日期范围，如果大于0，则只显示最近几天内更新的文章
'10       OrderType ---- 排序方式，1--按文章ID降序，2--按文章ID升序，3--按更新时间降序，4--按更新时间升序，5--按点击数降序，6--按点击数升序，7--按评论数降序，8--按评论数升序
'11       ShowType ---- 显示方式，1为普通样式，2为表格式，3为各项独立式，4为智能多列式，5为输出DIV，6为输出RSS
'12       TitleLen ---- 标题最多字符数，一个汉字=两个英文字符，若为0，则显示完整标题
'13       ContentLen ---- 文章内容最多字符数，一个汉字=两个英文字符，为0时不显示。请文章数量较多，可能会导致溢出错误。
'14       ShowClassName ---- 是否显示所属栏目名称，True为显示，False为不显示
'15       ShowPropertyType ---- 显示文章属性（固顶/推荐/普通）的方式，0为不显示，1为小图片，2为符号，3--9为小图片，10为序号
'16       ShowIncludePic ---- 是否显示“[图文]”字样，True为显示，False为不显示
'17       ShowAuthor ---- 是否显示文章作者，True为显示，False为不显示
'18       ShowDateType ---- 显示更新日期的样式，0为不显示，1为显示年月日，2为只显示月日，3为以“月-日”方式显示月日。
'19       ShowHits ---- 是否显示文章点击数，True为显示，False为不显示
'20       ShowHotSign ---- 是否显示热门文章标志，True为显示，False为不显示
'21       ShowNewSign ---- 是否显示新文章标志，True为显示，False为不显示
'22       ShowTips ---- 是否显示作者、更新日期、点击数等浮动提示信息，True为显示，False为不显示
'23       ShowCommentLink ---- 是否显示评论链接，True为显示，False为不显示，此选项只有当相应文章在后台设置了“显示评论链接”才有效。
'24       UsePage ---- 是否分页显示，True为分页显示，False为不分页显示，每页显示的文章数量由MaxPerPage指定
'25       OpenType ---- 文章打开方式，0为在原窗口打开，1为在新窗口打开
'26       Cols ---- 每行的列数。超过此列数就换行。
'27       CssNameA ---- 列表中文字链接调用的CSS类名
'28       CssName1 ---- 列表中奇数行的CSS效果的类名
'29       CssName2 ---- 列表中偶数行的CSS效果的类名
'=================================================
Public Function NewGetArticleList(iChannelID, arrClassID, IncludeChild, iSpecialID, UrlType, ArticleNum, IsHot, IsElite, Author, DateNum, OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowIncludePic, ShowAuthor, ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, ShowCommentLink, UsePage, OpenType, Cols, CssNameA, CssName1, CssName2)
    Dim sqlInfo, rsInfoList, strInfoList, CssName, iCount, iNumber, InfoUrl
    Dim strProperty, strTitle, strLink, strAuthor, strUpdateTime, strHits, strHotSign, strNewSign, strContent, strClassName
    Dim iTitleLen, strCommentLink
    Dim TDWidth_Author, TdWidth_Date

    TDWidth_Author = 10 * AuthorInfoLen
    TdWidth_Date = GetTdWidth_Date(ShowDateType)

    iCount = 0
    UrlType = PE_CLng(UrlType)
    Cols = PE_CLng1(Cols)

    If ShowType = 6 Then UrlType = 1
    If TitleLen < 0 Or TitleLen > 200 Then TitleLen = 50
    If IsNull(CssNameA) Then CssNameA = "listA"
    If IsNull(CssName1) Then CssName1 = "listbg"
    If IsNull(CssName2) Then CssName2 = "listbg2"

    FoundErr = False
    If (PE_Clng(iChannelID) <> 0 and Instr(iChannelID,",")=0) and (PE_Clng(iChannelID)<>PrevChannelID Or ChannelID = 0) Then
        Call GetChannel(PE_Clng(iChannelID))
        PrevChannelID = iChannelID       
    End If
    If FoundErr = True Then
        NewGetArticleList = ErrMsg
        Exit Function
    End If

    sqlInfo = "select"
    If ArticleNum > 0 Then
        If ShowType = 4 Then
            sqlInfo = sqlInfo & " top " & ArticleNum * 4
        Else
            sqlInfo = sqlInfo & " top " & ArticleNum
        End If
    End If
    sqlInfo = sqlInfo & " A.ChannelID,A.ClassID,A.ArticleID,A.Title,A.TitleFontColor,A.TitleFontType,A.ShowCommentLink,A.IncludePic,A.Author,A.UpdateTime,A.Hits,A.OnTop,A.Elite,A.InfoPurview,A.InfoPoint"
    If ContentLen > 0 Then
        sqlInfo = sqlInfo & ",A.Intro,A.Content"
    End If
    sqlInfo = sqlInfo & ",C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview"
    sqlInfo = sqlInfo & GetSqlStr(iChannelID, arrClassID, IncludeChild, iSpecialID, IsHot, IsElite, Author, DateNum, OrderType, ShowClassName, False)
    Set rsInfoList = Server.CreateObject("ADODB.Recordset")
    rsInfoList.Open sqlInfo, Conn, 1, 1
    If rsInfoList.BOF And rsInfoList.EOF Then
        If UsePage = True Then totalPut = 0
        If ShowType < 6 Then
            strInfoList = GetInfoList_StrNoItem(arrClassID, iSpecialID, IsHot, IsElite, strHot, strElite)
        End If
        rsInfoList.Close
        Set rsInfoList = Nothing
        NewGetArticleList = strInfoList
        Exit Function
    End If
    If UsePage = True And ShowType < 6 Then
        totalPut = rsInfoList.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                iMod = 0
                If CurrentPage > UpdatePages Then
                    iMod = totalPut Mod MaxPerPage
                    If iMod <> 0 Then iMod = MaxPerPage - iMod
                End If
                rsInfoList.Move (CurrentPage - 1) * MaxPerPage - iMod
            Else
                CurrentPage = 1
            End If
        End If
    End If

    CssName = CssName1

    If ShowType = 6 Then Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
    If ShowType = 2 Or ShowType = 4 Or (Cols > 1 and ShowType<>5) Then
        strInfoList = "<table width=""100%"" cellpadding=""0"" cellspacing=""0""><tr>"
    Else
        strInfoList = ""
    End If

    Dim CurrentTitleLen, isfirst, rownum, outend
    CurrentTitleLen = 0
    isfirst = True
    rownum = 1
    outend = False
    Do While Not rsInfoList.EOF
        'If iChannelID = 0 Then
            If rsInfoList("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsInfoList("ChannelID"))
                PrevChannelID = rsInfoList("ChannelID")
            End If
       ' End If
        If UsePage = True Then
            iNumber = (CurrentPage - 1) * MaxPerPage + iCount + 1
        Else
            iNumber = iCount + 1
        End If

        ChannelUrl = UrlPrefix(UrlType, ChannelUrl) & ChannelUrl
        ChannelUrl_ASPFile = UrlPrefix(UrlType, ChannelUrl_ASPFile) & ChannelUrl_ASPFile
        InfoUrl = GetArticleUrl(rsInfoList("ParentDir"), rsInfoList("ClassDir"), rsInfoList("UpdateTime"), rsInfoList("ArticleID"), rsInfoList("ClassPurview"), rsInfoList("InfoPurview"), rsInfoList("InfoPoint"))
        If ShowType < 6 And ShowType <> 4 Then

            strProperty = GetInfoList_GetStrProperty(ShowPropertyType, rsInfoList("OnTop"), rsInfoList("Elite"), iNumber, strCommon, strTop, strElite)
            strHotSign = GetInfoList_GetStrHotSign(ShowHotSign, rsInfoList("Hits"), strHot)
            strNewSign = GetInfoList_GetStrNewSign(ShowNewSign, rsInfoList("UpdateTime"), strNew)
            strCommentLink = GetInfoList_GetStrCommentLink(ShowCommentLink, rsInfoList("ShowCommentLink"), rsInfoList("ArticleID"))
            strAuthor = GetSubStr(rsInfoList("Author"), AuthorInfoLen, True)
            strUpdateTime = GetInfoList_GetStrUpdateTime(rsInfoList("UpdateTime"), ShowDateType)
            strHits = rsInfoList("Hits")
            If ShowType = 3 Or ShowType = 5 Then
                strAuthor = GetInfoList_GetStrAuthor_Xml(ShowAuthor, strAuthor)
                strUpdateTime = GetInfoList_GetStrUpdateTime_Xml(ShowDateType, strUpdateTime)
                strHits = GetInfoList_GetStrHits_Xml(ShowHits, strHits)
            End If

            iTitleLen = GetInfoList_GetTitleLen(TitleLen, ShowIncludePic, ShowCommentLink, rsInfoList("IncludePic"), rsInfoList("ShowCommentLink"))
            strTitle = GetInfoList_GetStrTitle(rsInfoList("Title"), iTitleLen, rsInfoList("TitleFontType"), rsInfoList("TitleFontColor"))

            strLink = ""
            If ShowClassName = True Then
                strLink = strLink & GetInfoList_GetStrClassLink(Character_Class, CssNameA, rsInfoList("ClassID"), rsInfoList("ClassName"), GetClassUrl(rsInfoList("ParentDir"), rsInfoList("ClassDir"), rsInfoList("ClassID"), rsInfoList("ClassPurview")))
            End If
            If ShowIncludePic = True Then
                strLink = strLink & GetInfoList_GetStrIncludePic(rsInfoList("IncludePic"))
            End If
            strLink = strLink & GetInfoList_GetStrInfoLink(strList_Title, ShowTips, OpenType, CssNameA, strTitle, InfoUrl, rsInfoList("Title"), rsInfoList("Author"), rsInfoList("UpdateTime"))
            strContent = ""
            Select Case PE_CLng(ShowType)
            Case 1, 3, 5
                If ContentLen > 0 Then
                    strContent = strContent & "<div " & strList_Content_Div & ">"
                    strContent = strContent & GetInfoList_GetStrContent(ContentLen, rsInfoList("Content"), rsInfoList("Intro"))
                    strContent = strContent & "</div>"
                End If
            Case 2
                If ContentLen > 0 Then
                    strContent = "<tr><td colspan=""10"" class=""" & CssName & """>"
                    strContent = strContent & GetInfoList_GetStrContent(ContentLen, rsInfoList("Content"), rsInfoList("Intro"))
                    strContent = strContent & "</td></tr>"
                End If
            End Select

        ElseIf ShowType = 6 Then
            strTitle = GetInfoList_GetStrTitle(rsInfoList("Title"), TitleLen, rsInfoList("TitleFontType"), rsInfoList("TitleFontColor"))
            strTitle = ReplaceText(xml_nohtml(strTitle), 2)
            strLink = InfoUrl
            If ContentLen > 0 Then
                If Trim(rsInfoList("Intro") & "") = "" Then
                    strContent = Left(Replace(Replace(Replace(xml_nohtml(rsInfoList("Content")), "[NextPage]", ""), ">", "&gt;"), "<", "&lt;"), ContentLen)
                Else
                    strContent = Left(xml_nohtml(rsInfoList("Intro")), ContentLen)
                End If
            End If
            strAuthor = GetInfoList_GetStrAuthor_RSS(Author)
            If ShowClassName = True And rsInfoList("ClassID") <> -1 Then
                strClassName = xml_nohtml(rsInfoList("ClassName"))
            Else
                strClassName = ""
            End If
            strUpdateTime = GetInfoList_GetStrUpdateTime(rsInfoList("UpdateTime"), ShowDateType)

        End If

        Select Case PE_CLng(ShowType)
        Case 1
            If Cols > 1 Then
                strInfoList = strInfoList & "<td valign=""top"" class=""" & CssName & """>"
            End If
            strInfoList = strInfoList & strProperty & "&nbsp;" & strLink
            strInfoList = strInfoList & GetInfoList_GetStrAuthorDateHits(ShowAuthor, ShowDateType, ShowHits, strAuthor, strUpdateTime, strHits, rsInfoList("ChannelID"))
            strInfoList = strInfoList & strHotSign & strNewSign & strCommentLink & strContent & "<br />"

            iCount = iCount + 1
            If Cols > 1 Then
                strInfoList = strInfoList & "</td>"
                If iCount Mod Cols = 0 Then
                    strInfoList = strInfoList & "</tr><tr>"
                    If iCount Mod 2 = 0 Then
                        CssName = CssName1
                    Else
                        CssName = CssName2
                    End If
                End If
            End If
        Case 2
            If strProperty <> "" Then
                strInfoList = strInfoList & "<td width=""10"" valign=""top"" class=""" & CssName & """>" & strProperty & "</td>"
            End If
            strInfoList = strInfoList & "<td class=""" & CssName & """>" & strLink & strHotSign & strNewSign & strCommentLink & "</td>"
            If ShowAuthor = True Then
                strInfoList = strInfoList & "<td align=""center"" class=""" & CssName & """ width=""" & TDWidth_Author & """>" & strAuthor & "</td>"
            End If
            If ShowDateType > 0 Then
                strInfoList = strInfoList & "<td align=""right"" class=""" & CssName & """ width=""" & TdWidth_Date & """>" & strUpdateTime & "</td>"
            End If
            If ShowHits = True Then
                strInfoList = strInfoList & "<td align=""center"" class=""" & CssName & """ width=""40"">" & strHits & "</td>"
            End If

            iCount = iCount + 1
            If (iCount Mod Cols = 0) Or ContentLen > 0 Then
                strInfoList = strInfoList & "</tr>"
                strInfoList = strInfoList & strContent
                strInfoList = strInfoList & "<tr>"
                If iCount Mod (Cols * 2) = 0 Then
                    CssName = CssName1
                Else
                    CssName = CssName2
                End If
            End If
        Case 3
            If Cols > 1 Then
                strInfoList = strInfoList & "<td valign=""top"" class=""" & CssName & """>"
            End If
            strInfoList = strInfoList & strProperty & "&nbsp;" & strLink
            strInfoList = strInfoList & strAuthor & strUpdateTime & strHits
            strInfoList = strInfoList & strHotSign & strNewSign & strCommentLink & strContent
            strInfoList = strInfoList & "<br />"

            iCount = iCount + 1
            If Cols > 1 Then
                strInfoList = strInfoList & "</td>"
                If iCount Mod Cols = 0 Then
                    strInfoList = strInfoList & "</tr><tr>"
                    If iCount Mod 2 = 0 Then
                        CssName = CssName1
                    Else
                        CssName = CssName2
                    End If
                End If
            End If
        Case 5 '输出DIV
            strInfoList = strInfoList & "<ul class=""" & CssName & """>"
            strInfoList = strInfoList & strProperty & strLink
            strInfoList = strInfoList & strAuthor & "<span>" & strUpdateTime & "</span>" & strHits
            strInfoList = strInfoList & strHotSign & strNewSign & strCommentLink & strContent
            strInfoList = strInfoList & "</ul>"

            iCount = iCount + 1
            If iCount Mod 2 = 0 Then
                CssName = CssName1
            Else
                CssName = CssName2
            End If
        Case 6 '输出RSS
            strInfoList = strInfoList & GetInfoList_GetStrRSS(strTitle, strLink, strContent, strAuthor, strClassName, strUpdateTime)
            iCount = iCount + 1
        Case 4 '输出智能多列式
            If TitleLen > 0 Then
                strTitle = ReplaceText(GetSubStr(rsInfoList("Title"), TitleLen, ShowSuspensionPoints), 2)
            Else
                strTitle = ReplaceText(rsInfoList("Title"), 2)
            End If
            iTitleLen = Charlong(strTitle)
            CurrentTitleLen = CurrentTitleLen + iTitleLen

            strLink = ""
            strLink = strLink & GetInfoList_GetStrInfoLink(strList_Title, ShowTips, OpenType, CssNameA, strTitle, InfoUrl, rsInfoList("Title"), rsInfoList("Author"), rsInfoList("UpdateTime"))
             
            If ShowCommentLink = True And rsInfoList("ShowCommentLink") = True Then
                strLink = strLink & "&nbsp;<a href='" & ChannelUrl_ASPFile & "/Comment.asp?Action=ShowAll&ArticleID=" & rsInfoList("ArticleID") & "'>" & strComment & "</a>"
                CurrentTitleLen = CurrentTitleLen + 1 + Charlong(nohtml(strComment))
            End If
             
            If isfirst = True Then
                strInfoList = strInfoList & "<td valign='top' class='" & CssName & "'>" & strProperty & strLink
                rownum = rownum + 1
                If CurrentTitleLen > TitleLen + 1 Then
                    CurrentTitleLen = 0
                    If rownum > ArticleNum Then
                        strInfoList = strInfoList & "</td></tr>"
                        Exit Do
                    Else
                        strInfoList = strInfoList & "</td></tr><tr>"
                    End If
                    iCount = iCount + 1
                Else
                    isfirst = False
                    CurrentTitleLen = CurrentTitleLen + 1
                End If
                If iCount Mod 2 = 0 Then
                    CssName = CssName1
                Else
                    CssName = CssName2
                End If
            Else
                If CurrentTitleLen > TitleLen + 1 And outend = False Then
                    CurrentTitleLen = iTitleLen
                    If ShowCommentLink = True And rsInfoList("ShowCommentLink") = True Then
                        CurrentTitleLen = CurrentTitleLen + 1 + Charlong(nohtml(strComment))
                    End If
             
                    strInfoList = strInfoList & "</td></tr><tr>"
                    iCount = iCount + 1
                    If iCount Mod 2 = 0 Then
                        CssName = CssName1
                    Else
                        CssName = CssName2
                    End If
                    strInfoList = strInfoList & "<td valign='top' class='" & CssName & "'>" & strProperty & strLink
                    rownum = rownum + 1
                    If rownum > ArticleNum Then
                        If CurrentTitleLen >= TitleLen Then
                            strInfoList = strInfoList & "</td></tr>"
                            Exit Do
                        Else
                            outend = True
                        End If
                    End If
                Else
                    If CurrentTitleLen > TitleLen + 1 Then
                        strInfoList = strInfoList & "</td></tr>"
                        Exit Do
                    Else
                        strInfoList = strInfoList & "&nbsp;" & strLink
                        CurrentTitleLen = CurrentTitleLen + 1
                    End If
                End If
            End If
        End Select
        rsInfoList.MoveNext
        If UsePage = True And iCount >= MaxPerPage Then Exit Do
    Loop
    If ShowType = 4 Then
        strInfoList = strInfoList & "</table>"  
    ElseIF ShowType = 2 Or (Cols > 1 and ShowType<>5) Then
        strInfoList = strInfoList & "</tr></table>"
    End If

    rsInfoList.Close
    Set rsInfoList = Nothing
    If ShowType = 6 And RssCodeType = False Then strInfoList = unicode(strInfoList)
    NewGetArticleList = strInfoList
End Function
%>
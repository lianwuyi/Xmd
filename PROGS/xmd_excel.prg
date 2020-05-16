*打印

set cent on
set date to ansi
set talk off
SET SAFETY OFF
SET DELETE ON
SET EXCLUSIVE OFF
close data all
close table all

* 打印

*----- 
#INCLUDE Excel8.h
#DEFINE False .F.
#DEFINE True .T.
LOCAL loExcel, lcOldError, lcRange, lnSheets, lnCounter
WAIT WINDOW  "正在收集数据......" NOWAIT NOCLEAR

**
    SELECT * ;
        FROM ..\data\xmd.dbf  ;
        ORDER BY 码单id ASC ;
        INTO CURSOR Output 
               
*        ORDER BY 单据号 ASC ; 

wait window '正在启动 ‘EXCEL’表格，请稍候……' nowait noclear


* 创建EXCEL 对象，添加EXCEL模版
lcOldError = ON("ERROR")
ON ERROR loExcel = .NULL.
loExcel = GetObject(, "Excel.Application")
ON ERROR &lcOldError

IF ISNULL(loExcel)
     loExcel = CreateObject( "Excel.Application" )
ENDIF

loExcel.visible = .f.                                 && 让 EXCEL 可视 / .f. 为不可视

* .ActiveWorkbook.Close  && 关闭打开的文件

loExcel.workbooks.add('c:\WINDOWS\Application Data\Microsoft\Templates\联业细码单.xlt')   
* && 添加模版,模版默认保存位置： 'C:\WINDOWS\Application Data\Microsoft\Templates'

WAIT WINDOW "正在写入 Excel 电子表格数据，请等候......" NOWAIT NOCLEAR
*
mdids = alltrim(str(码单id))
if len(mdids) = 9
   mdids = mdids
endif
if len(mdids) = 8
   mdids = "0"+mdids
endif
if len(mdids) = 7
   mdids = "00"+mdids
endif
if len(mdids) = 6
   mdids = "000"+mdids
endif
if len(mdids) = 5
   mdids = "0000"+mdids
endif
if len(mdids) = 4
   mdids = "00000"+mdids
endif
if len(mdids) = 3
   mdids = "000000"+mdids
endif
if len(mdids) = 2
   mdids = "0000000"+mdids
endif
if len(mdids) = 1
   mdids = "00000000"+mdids
endif

*
WITH loExcel
  ****
     WITH .Range("R3")
          .Value = mdids
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****
  ****
     WITH .Range("C4")
          .Value = ALLTRIM(客户名称)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****
  ****
     WITH .Range("R4")
          .Value = dtoc(出货日期)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****
****
     WITH .Range("C5")
          .Value = ALLTRIM(货品编码)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****
  ****
     WITH .Range("K5")
          .Value = ALLTRIM(付款期限天)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****
  ****
     WITH .Range("P5")
          .Value = ALLTRIM(客户订单号)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****


  ****
     WITH .Range("C6")
          .Value = ALLTRIM(布别)+ALLTRIM(布幅)+ALLTRIM(组织)+"/"+ alltrim(支数) +"   颜色："+ ALLTRIM(颜色)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****
  ****
     WITH .Range("P6")
          .Value = ALLTRIM(客户加工号)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****
  ****
     WITH .Range("C7")
          .Value = ALLTRIM(交货地点)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH   
****
  ****
     WITH .Range("P7")
          .Value = ALLTRIM(客户跟单员)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
  ****
     WITH .Range("A8")
          .Value = ALLTRIM(表内单位)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH      
  ****
     WITH .Range("B8")
          .Value = ALLTRIM(颜色1)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH  

*****第一行 a
**********
if a1<>0
     WITH .Range("B10")
          .Value = a1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa1))>0
     WITH .Range("C10")
          .Value = aa1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a2<>0
     WITH .Range("D10")
          .Value = a2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa2))>0
     WITH .Range("E10")
          .Value = aa2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a3<>0
     WITH .Range("F10")
          .Value = a3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa3))>0
     WITH .Range("G10")
          .Value = aa3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a4<>0
     WITH .Range("H10")
          .Value = a4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa4))>0
     WITH .Range("I10")
          .Value = aa4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a5<>0
     WITH .Range("J10")
          .Value = a5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa5))>0
     WITH .Range("K10")
          .Value = aa5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a6<>0
     WITH .Range("L10")
          .Value = a6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa6))>0
     WITH .Range("M10")
          .Value = aa6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a7<>0
     WITH .Range("N10")
          .Value = a7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa7))>0
     WITH .Range("O10")
          .Value = aa7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a8<>0
     WITH .Range("P10")
          .Value = a8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa8))>0
     WITH .Range("Q10")
          .Value = aa8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a9<>0
     WITH .Range("R10")
          .Value = a9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa9))>0
     WITH .Range("S10")
          .Value = aa9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if a10<>0
     WITH .Range("T10")
          .Value = a10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(aa10))>0
     WITH .Range("U10")
          .Value = aa10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
* 第二行 b, 
**********
if b1<>0
     WITH .Range("B11")
          .Value = b1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb1))>0
     WITH .Range("C11")
          .Value = bb1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b2<>0
     WITH .Range("D11")
          .Value = b2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb2))>0
     WITH .Range("E11")
          .Value = bb2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b3<>0
     WITH .Range("F11")
          .Value = b3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb3))>0
     WITH .Range("G11")
          .Value = bb3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b4<>0
     WITH .Range("H11")
          .Value = b4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb4))>0
     WITH .Range("I11")
          .Value = bb4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b5<>0
     WITH .Range("J11")
          .Value = b5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb5))>0
     WITH .Range("K11")
          .Value = bb5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b6<>0
     WITH .Range("L11")
          .Value = b6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb6))>0
     WITH .Range("M11")
          .Value = bb6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b7<>0
     WITH .Range("N11")
          .Value = b7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb7))>0
     WITH .Range("O11")
          .Value = bb7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b8<>0
     WITH .Range("P11")
          .Value = b8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb8))>0
     WITH .Range("Q11")
          .Value = bb8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b9<>0
     WITH .Range("R11")
          .Value = b9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb9))>0
     WITH .Range("S11")
          .Value = bb9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if b10<>0
     WITH .Range("T11")
          .Value = b10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(bb10))>0
     WITH .Range("U11")
          .Value = bb10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************

*** 第三行 c
if c1<>0
     WITH .Range("B12")
          .Value = c1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc1))>0
     WITH .Range("C12")
          .Value = cc1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c2<>0
     WITH .Range("D12")
          .Value = c2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc2))>0
     WITH .Range("E12")
          .Value = cc2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c3<>0
     WITH .Range("F12")
          .Value = c3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc3))>0
     WITH .Range("G12")
          .Value = cc3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c4<>0
     WITH .Range("H12")
          .Value = c4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc4))>0
     WITH .Range("I12")
          .Value = cc4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c5<>0
     WITH .Range("J12")
          .Value = c5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc5))>0
     WITH .Range("K12")
          .Value = cc5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c6<>0
     WITH .Range("L12")
          .Value = c6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc6))>0
     WITH .Range("M12")
          .Value = cc6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c7<>0
     WITH .Range("N12")
          .Value = c7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc7))>0
     WITH .Range("O12")
          .Value = cc7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c8<>0
     WITH .Range("P12")
          .Value = c8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc8))>0
     WITH .Range("Q12")
          .Value = cc8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c9<>0
     WITH .Range("R12")
          .Value = c9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc9))>0
     WITH .Range("S12")
          .Value = cc9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if c10<>0
     WITH .Range("T12")
          .Value = c10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(cc10))>0
     WITH .Range("U12")
          .Value = cc10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
** 第四行 d

if d1<>0
     WITH .Range("B13")
          .Value = d1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd1))>0
     WITH .Range("C13")
          .Value = dd1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d2<>0
     WITH .Range("D13")
          .Value = d2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd2))>0
     WITH .Range("E13")
          .Value = dd2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d3<>0
     WITH .Range("F13")
          .Value = d3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd3))>0
     WITH .Range("G13")
          .Value = dd3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d4<>0
     WITH .Range("H13")
          .Value = d4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd4))>0
     WITH .Range("I13")
          .Value = dd4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d5<>0
     WITH .Range("J13")
          .Value = d5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd5))>0
     WITH .Range("K13")
          .Value = dd5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d6<>0
     WITH .Range("L13")
          .Value = d6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd6))>0
     WITH .Range("M13")
          .Value = dd6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d7<>0
     WITH .Range("N13")
          .Value = d7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd7))>0
     WITH .Range("O13")
          .Value = dd7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d8<>0
     WITH .Range("P13")
          .Value = d8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd8))>0
     WITH .Range("Q13")
          .Value = dd8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d9<>0
     WITH .Range("R13")
          .Value = d9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd9))>0
     WITH .Range("S13")
          .Value = dd9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if d10<>0
     WITH .Range("T13")
          .Value = d10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(dd10))>0
     WITH .Range("U13")
          .Value = dd10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif

** 第五行 E
if e1<>0
     WITH .Range("B14")
          .Value = e1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee1))>0
     WITH .Range("C14")
          .Value = ee1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e2<>0
     WITH .Range("D14")
          .Value = e2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee2))>0
     WITH .Range("E14")
          .Value = ee2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e3<>0
     WITH .Range("F14")
          .Value = e3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee3))>0
     WITH .Range("G14")
          .Value = ee3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e4<>0
     WITH .Range("H14")
          .Value = e4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee4))>0
     WITH .Range("I14")
          .Value = ee4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e5<>0
     WITH .Range("J14")
          .Value = e5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee5))>0
     WITH .Range("K14")
          .Value = ee5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e6<>0
     WITH .Range("L14")
          .Value = e6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee6))>0
     WITH .Range("M14")
          .Value = ee6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e7<>0
     WITH .Range("N14")
          .Value = e7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee7))>0
     WITH .Range("O14")
          .Value = ee7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e8<>0
     WITH .Range("P14")
          .Value = e8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee8))>0
     WITH .Range("Q14")
          .Value = ee8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e9<>0
     WITH .Range("R14")
          .Value = e9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee9))>0
     WITH .Range("S14")
          .Value = ee9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if e10<>0
     WITH .Range("T14")
          .Value = e10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ee10))>0
     WITH .Range("U14")
          .Value = ee10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif


** 第六行 f
if f1<>0
     WITH .Range("B15")
          .Value = f1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff1))>0
     WITH .Range("C15")
          .Value = ff1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f2<>0
     WITH .Range("D15")
          .Value = f2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff2))>0
     WITH .Range("E15")
          .Value = ff2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f3<>0
     WITH .Range("F15")
          .Value = f3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff3))>0
     WITH .Range("G15")
          .Value = ff3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f4<>0
     WITH .Range("H15")
          .Value = f4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff4))>0
     WITH .Range("I15")
          .Value = ff4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f5<>0
     WITH .Range("J15")
          .Value = f5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff5))>0
     WITH .Range("K15")
          .Value = ff5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f6<>0
     WITH .Range("L15")
          .Value = f6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff6))>0
     WITH .Range("M15")
          .Value = ff6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f7<>0
     WITH .Range("N15")
          .Value = f7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff7))>0
     WITH .Range("O15")
          .Value = ff7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f8<>0
     WITH .Range("P15")
          .Value = f8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff8))>0
     WITH .Range("Q15")
          .Value = ff8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f9<>0
     WITH .Range("R15")
          .Value = f9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff9))>0
     WITH .Range("S15")
          .Value = ff9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if f10<>0
     WITH .Range("T15")
          .Value = f10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ff10))>0
     WITH .Range("U15")
          .Value = ff10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif

** 第七行 G
if g1<>0
     WITH .Range("B16")
          .Value = g1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg1))>0
     WITH .Range("C16")
          .Value = gg1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g2<>0
     WITH .Range("D16")
          .Value = g2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg2))>0
     WITH .Range("E16")
          .Value = gg2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g3<>0
     WITH .Range("F16")
          .Value = g3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg3))>0
     WITH .Range("G16")
          .Value = gg3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g4<>0
     WITH .Range("H16")
          .Value = g4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg4))>0
     WITH .Range("I16")
          .Value = gg4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g5<>0
     WITH .Range("J16")
          .Value = g5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg5))>0
     WITH .Range("K16")
          .Value = gg5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g6<>0
     WITH .Range("L16")
          .Value = g6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg6))>0
     WITH .Range("M16")
          .Value = gg6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g7<>0
     WITH .Range("N16")
          .Value = g7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg7))>0
     WITH .Range("O16")
          .Value = gg7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g8<>0
     WITH .Range("P16")
          .Value = g8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg8))>0
     WITH .Range("Q16")
          .Value = gg8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g9<>0
     WITH .Range("R16")
          .Value = g9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg9))>0
     WITH .Range("S16")
          .Value = gg9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if g10<>0
     WITH .Range("T16")
          .Value = g10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(gg10))>0
     WITH .Range("U16")
          .Value = gg10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif

** 第八行 h
if h1<>0
     WITH .Range("B17")
          .Value = h1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh1))>0
     WITH .Range("C17")
          .Value = hh1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h2<>0
     WITH .Range("D17")
          .Value = h2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh2))>0
     WITH .Range("E17")
          .Value = hh2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h3<>0
     WITH .Range("F17")
          .Value = h3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh3))>0
     WITH .Range("G17")
          .Value = hh3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h4<>0
     WITH .Range("H17")
          .Value = h4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh4))>0
     WITH .Range("I17")
          .Value = hh4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h5<>0
     WITH .Range("J17")
          .Value = h5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh5))>0
     WITH .Range("K17")
          .Value = hh5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h6<>0
     WITH .Range("L17")
          .Value = h6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh6))>0
     WITH .Range("M17")
          .Value = hh6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h7<>0
     WITH .Range("N17")
          .Value = h7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh7))>0
     WITH .Range("O17")
          .Value = hh7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h8<>0
     WITH .Range("P17")
          .Value = h8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh8))>0
     WITH .Range("Q17")
          .Value = hh8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h9<>0
     WITH .Range("R17")
          .Value = h9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh9))>0
     WITH .Range("S17")
          .Value = hh9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if h10<>0
     WITH .Range("T17")
          .Value = h10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(hh10))>0
     WITH .Range("U17")
          .Value = hh10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif

***第九行 i
if i1<>0
     WITH .Range("B18")
          .Value = i1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii1))>0
     WITH .Range("C18")
          .Value = ii1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i2<>0
     WITH .Range("D18")
          .Value = i2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii2))>0
     WITH .Range("E18")
          .Value = ii2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i3<>0
     WITH .Range("F18")
          .Value = i3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii3))>0
     WITH .Range("G18")
          .Value = ii3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i4<>0
     WITH .Range("H18")
          .Value = i4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii4))>0
     WITH .Range("I18")
          .Value = ii4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i5<>0
     WITH .Range("J18")
          .Value = i5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii5))>0
     WITH .Range("K18")
          .Value = ii5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i6<>0
     WITH .Range("L18")
          .Value = i6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii6))>0
     WITH .Range("M18")
          .Value = ii6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i7<>0
     WITH .Range("N18")
          .Value = i7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii7))>0
     WITH .Range("O18")
          .Value = ii7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i8<>0
     WITH .Range("P18")
          .Value = i8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii8))>0
     WITH .Range("Q18")
          .Value = ii8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i9<>0
     WITH .Range("R18")
          .Value = i9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii9))>0
     WITH .Range("S18")
          .Value = ii9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if i10<>0
     WITH .Range("T18")
          .Value = i10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(ii10))>0
     WITH .Range("U18")
          .Value = ii10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
*** 第十行 j
if j1<>0
     WITH .Range("B19")
          .Value = j1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj1))>0
     WITH .Range("C19")
          .Value = jj1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j2<>0
     WITH .Range("D19")
          .Value = j2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj2))>0
     WITH .Range("E19")
          .Value = jj2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j3<>0
     WITH .Range("F19")
          .Value = j3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj3))>0
     WITH .Range("G19")
          .Value = jj3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j4<>0
     WITH .Range("H19")
          .Value = j4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj4))>0
     WITH .Range("I19")
          .Value = jj4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j5<>0
     WITH .Range("J19")
          .Value = j5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj5))>0
     WITH .Range("K19")
          .Value = jj5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j6<>0
     WITH .Range("L19")
          .Value = j6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj6))>0
     WITH .Range("M19")
          .Value = jj6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j7<>0
     WITH .Range("N19")
          .Value = j7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj7))>0
     WITH .Range("O19")
          .Value = jj7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j8<>0
     WITH .Range("P19")
          .Value = j8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj8))>0
     WITH .Range("Q19")
          .Value = jj8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j9<>0
     WITH .Range("R19")
          .Value = j9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj9))>0
     WITH .Range("S19")
          .Value = jj9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
************
**********
if j10<>0
     WITH .Range("T19")
          .Value = j10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif  
*
if len(alltrim(jj10))>0
     WITH .Range("U19")
          .Value = jj10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
**********
*

naj1 = a1+b1+c1+d1+e1+f1+g1+h1+i1+j1

naj2 = a2+b2+c2+d2+e2+f2+g2+h2+i2+j2

naj3 = a3+b3+c3+d3+e3+f3+g3+h3+i3+j3

naj4 = a4+b4+c4+d4+e4+f4+g4+h4+i4+j4

naj5 = a5+b5+c5+d5+e5+f5+g5+h5+i5+j5

naj6 = a6+b6+c6+d6+e6+f6+g6+h6+i6+j6

naj7 = a7+b7+c7+d7+e7+f7+g7+h7+i7+j7

naj8 =a8+b8+c8+d8+e8+f8+g8+h8+i8+j8

naj9 = a9+b9+c9+d9+e9+f9+g9+h9+i9+j9

naj10 = a10+b10+c10+d10+e10+f10+g10+h10+i10+j10
* 米数
* 码数
* 合计
**

     WITH .Range("A20")
          .Value = "小计("+ALLTRIM(表内单位)+")"
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 10
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 


**
if naj1>0
     WITH .Range("B20")
          .Value = naj1
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
*
if naj2>0
     WITH .Range("D20")
          .Value = naj2
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
*
if naj3>0
     WITH .Range("F20")
          .Value = naj3
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
*
if naj4 > 0
     WITH .Range("H20")
          .Value = naj4
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
*
if naj5>0
     WITH .Range("J20")
          .Value = naj5
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
*
if naj6>0
     WITH .Range("L20")
          .Value = naj6
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
*
if naj7>0
     WITH .Range("N20")
          .Value = naj7
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
if naj8>0
     WITH .Range("P20")
          .Value = naj8
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
if naj9>0
     WITH .Range("R20")
          .Value = naj9
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
if naj10>0
     WITH .Range("T20")
          .Value = naj10
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
* 合计

* 码数折米数  1 码 = 0.9144 米
* 米数折码数  1 米 = 1.0936 码
* 码数

***
if 疋数 > 0
     WITH .Range("B21")
          .Value = 疋数
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
**
if 米数 > 0
     WITH .Range("E21")
          .Value = 米数
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
**
**
if 码数 > 0
     WITH .Range("K21")
          .Value = 码数
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
endif
**
     WITH .Range("B22")
          .Value = ALLTRIM(备注)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 

     WITH .Range("B4")
          .Value = ALLTRIM(订单号)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
     
*
     WITH .Range("B25")
          .Value = ALLTRIM(组别)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
     WITH .Range("G25")
          .Value = ALLTRIM(装车人)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
     WITH .Range("B26")
          .Value = ALLTRIM(车号) 
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 
     WITH .Range("G26")
          .Value = ALLTRIM(制单员)
          WITH .font
               .NAME = "宋体"        && "Arial"
               .Bold = .f.    && T 粗体  ; F 正常
               .Size = 12
               .Underline = xlUnderlineStyleNone         &&xlUnderlineStyleSingle
             
          ENDWITH
     ENDWITH 

* 替换 颜色，组织 里面的‘*’号为‘X’号,去掉‘\’号，防止特殊字符产生的错误
   zz1 = CHRTRAN(ALLTRIM(组织),'*','X') 
   zz1 = CHRTRAN(ALLTRIM(zz1),'\','')      
   ys1 = CHRTRAN(ALLTRIM(颜色),'*',' ')
   ys1 = CHRTRAN(ALLTRIM(ys1),'\',' ')    
   
* 文件名称 
wjmc= ALLTRIM(客户名称)+'（'+ys1+zz1+'）'+ALLTRIM(STR(码单id))

loExcel.DisplayAlerts=.f.                   &&去除保存是弹出是否要覆盖的对话框
loExcel.visible = .T.                       && 让 EXCEL 可视 / .f. 为不可视
.ActiveWorkbook.SaveAs('d:\待发邮件\'+wjmc) && 保存,wjmc 为文件名称变量
.ActiveWorkbook.Close                       && 关闭打开的文件
******
* .ActiveWindow.SelectedSheets.PrintOut.Copies = 0
*  .ActiveWindow.SelectedSheets.PrintOut    && 打印
*   .ActiveWindow.SelectedSheets.PrintOut Copies:=1, Preview:=True

*  GETPRINTER()  && 设置打印机
  
*  .Preview=True
******
  Release loExcel
  CLOSE DATA ALL
  CLOSE TABLE ALL
* wait clear
wait window '操作成功：已生成EXCEL文件，保存在D：\待发邮件\文件夹中！' nowait noclear
close all
DO FORM ..\FORMS\4ck细码单查看1.scx

RETURN
ENDWITH

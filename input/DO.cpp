// DO.cpp : 实现文件
//

#include "stdafx.h"
#include "input.h"
#include "DO.h"
#include "afxdialogex.h"
#define WM_UPDATE_MESSAGE2 10004
#define WM_UPDATE_MESSAGE 10003

// DO 对话框
extern CString exfilename;//外部变量
extern CString m_inputopname;
extern BOOL m_closeornot;


IMPLEMENT_DYNAMIC(DO, CDialogEx)

DO::DO(CWnd* pParent /*=NULL*/)
	: CDialogEx(DO::IDD, pParent)
{

}

DO::~DO()
{
}

void DO::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Text(pDX, IDC_EDIT1, m_boardnum);
	DDX_Text(pDX, IDC_EDIT2, m_opname);

	DDX_Text(pDX, IDC_EDIT3, m_tempture);
	DDX_Text(pDX, IDC_EDIT5, m_humidity);

	DDX_Text(pDX, IDC_EDIT4, m_multimeter);
	DDX_Text(pDX, IDC_EDIT7, m_constant_volt_source);

	DDX_Text(pDX, IDC_EDIT10, m_signal_source);
	DDX_Text(pDX, IDC_EDIT11, m_测试工装);

	
	DDX_Text(pDX, IDC_EDIT12, m_edit12);
	DDX_Text(pDX, IDC_EDIT13, m_edit13);
	DDX_Text(pDX, IDC_EDIT14, m_edit14);
	DDX_Text(pDX, IDC_EDIT15, m_edit15);

	DDX_Text(pDX, IDC_EDIT16, m_edit16);
	DDX_Text(pDX, IDC_EDIT17, m_edit17);
	DDX_Text(pDX, IDC_EDIT18, m_edit18);
	DDX_Text(pDX, IDC_EDIT19, m_edit19);


}


BEGIN_MESSAGE_MAP(DO, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &DO::OnBnClickedButton1)
	ON_MESSAGE(WM_UPDATE_MESSAGE, OnUpdateTrueMessage)
	ON_MESSAGE(WM_UPDATE_MESSAGE2, OnUpdateFalseMessage)
	ON_BN_CLICKED(IDC_BUTTON2, &DO::OnBnClickedButton2)
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_BUTTON10, &DO::OnBnClickedButton10)
	ON_BN_CLICKED(IDC_BUTTON11, &DO::OnBnClickedButton11)
	ON_BN_CLICKED(IDC_BUTTON12, &DO::OnBnClickedButton12)
	ON_BN_CLICKED(IDCANCEL, &DO::OnBnClickedCancel)
END_MESSAGE_MAP()


// DO 消息处理程序
LRESULT DO::OnUpdateTrueMessage(WPARAM wParam, LPARAM lParam)
{
	BOOL f=(BOOL)wParam;
	if(f==TRUE)
       UpdateData(TRUE);
	else
	   UpdateData(FALSE);
    return 0;
}
LRESULT DO::OnUpdateFalseMessage(WPARAM wParam, LPARAM lParam)
{
    UpdateData(FALSE);
    return 0;
}



void DO::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码

	if(m_click==TRUE)
	{
		m_click=FALSE;
	    AfxBeginThread((AFX_THREADPROC)DO::Button1Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("文件未关闭，不允许操作，请先关闭文件！"));
	}
}

DWORD DO::Button1Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	DO *pMain=(DO*) lpParam;	

	std::vector<CString> vec_str;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!(pMain->app.CreateDispatch(L"Excel.Application")))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);

	
	
	//得到Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//得到Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)1));

	


	pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);

	vec_str.push_back(_T("D"));
	vec_str.push_back(_T("DO"));
	vec_str.push_back(pMain->m_boardnum);
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(_T("合格（  √  ） /  不合格（   ）"));
	vec_str.push_back(pMain->m_tempture);
	vec_str.push_back(pMain->m_humidity);

	vec_str.push_back(pMain->m_multimeter);
	vec_str.push_back(pMain->m_constant_volt_source);
	
	vec_str.push_back(pMain->m_测试工装);
	vec_str.push_back(pMain->m_edit19);
	

	vec_Pos.push_back(_T("B3"));
	vec_Pos.push_back(_T("B4"));
    vec_Pos.push_back(_T("B5"));
	vec_Pos.push_back(_T("B6"));
    vec_Pos.push_back(_T("B7"));

    vec_Pos.push_back(_T("B9"));
    vec_Pos.push_back(_T("B10"));
    vec_Pos.push_back(_T("B11"));

	vec_Pos.push_back(_T("D13"));
    vec_Pos.push_back(_T("D14"));
    vec_Pos.push_back(_T("D15"));
	vec_Pos.push_back(_T("B3"));

    

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	//pMain->app.put_Visible(TRUE);
	//pMain->book.Save();
	//pMain->book.put_Saved(TRUE);
	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)2));


	vec_str.clear();
	vec_Pos.clear();

	for(size_t i=0;i<12;i++)
	{
		vec_str.push_back(_T("√"));
	}
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(_T("合格√ 不合格"));

	vec_Pos.push_back(_T("F3"));
	vec_Pos.push_back(_T("F4"));
    vec_Pos.push_back(_T("F5"));
	vec_Pos.push_back(_T("F6"));
    vec_Pos.push_back(_T("F7"));
	vec_Pos.push_back(_T("F8"));
    vec_Pos.push_back(_T("F9"));
	vec_Pos.push_back(_T("F10"));
    vec_Pos.push_back(_T("F11"));
    vec_Pos.push_back(_T("F12"));
	vec_Pos.push_back(_T("F13"));
    vec_Pos.push_back(_T("F14"));

    vec_Pos.push_back(_T("B15"));
	vec_Pos.push_back(_T("D15"));
    vec_Pos.push_back(_T("G15"));

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}


	

	pMain->books.Close(); 
	
    pMain->app.Quit();  			// 退出
	//释放对象  
	pMain->range.ReleaseDispatch();
	pMain->sheet.ReleaseDispatch();
	pMain->sheets.ReleaseDispatch();
	pMain->book.ReleaseDispatch();
	pMain->books.ReleaseDispatch();
	pMain->app.ReleaseDispatch();


	CoUninitialize(); 

	pMain->m_click=TRUE;

	return TRUE;
}



BOOL DO::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化

	 m_boardnum=_T("DO_201901001");
	 m_opname=_T("路人甲");
	 m_tempture=_T("25℃");
	 m_humidity=_T("40%");

	 m_multimeter=_T("20190620-20200619");
	 m_constant_volt_source=_T("20190620-20200619");
	 m_signal_source=_T("20190620-20200619");
	 m_测试工装=_T("20190620-20200619");

	 m_OKorNot=_T("合格√ 不合格");


	 m_edit12=_T("40,12,8");
	 m_edit13=_T("8,8,8,8");
	 m_edit16=_T("8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8");

	 m_edit19=_T("单板环境试验后");




	 my_Font.CreatePointFont(140, L"Times NewRoman");//创建字体和大小

	 m_Brush.CreateSolidBrush(RGB(53,203,60));

	 GetDlgItem(IDC_EDIT12)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT13)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT14)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT15)->SetFont(&my_Font);

	 GetDlgItem(IDC_EDIT16)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT17)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT18)->SetFont(&my_Font);


	  m_click=TRUE;//初始可以单击


	  COleDateTime t = COleDateTime::GetCurrentTime();

	  m_strtime = t.Format(_T("%Y/%m/%d"));//打印填表日期




	 UpdateData(FALSE);

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void DO::OnBnClickedButton2()
{
	// TODO: 在此添加控件通知处理程序代码
	if(m_click==TRUE)
	{
		m_click=FALSE;
	    AfxBeginThread((AFX_THREADPROC)DO::Button2Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("文件未关闭，不允许操作，请先关闭文件！"));
	}


}


std::vector<float> DO::CStringtoFloat(CString m_s)
{
	int count=m_s.Replace(',',' ');
	std::vector<float> retout;

	if(count<=0)
	{
		//AfxMessageBox(_T("No Data"));
		return retout;
	}

	int pos=m_s.Find(' ');
	int i=0;

	while(pos!=-1)
	{
		CString field=m_s.Left(pos);
		retout.push_back(_tstof((const wchar_t*)field.GetBuffer(0)));
		i++;

		m_s=m_s.Right(m_s.GetLength()-pos-1);

	    pos=m_s.Find(' ');
	}

		if(m_s.GetLength()>0)
		{
			retout.push_back(_tstof((const wchar_t*)m_s.GetBuffer(0)));
		}

	

      return retout;
}


DWORD DO::Button2Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	DO *pMain=(DO*) lpParam;	

	pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);

	m_inputopname=pMain->m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!pMain->app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//得到Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//得到Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)3));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_edit12);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("输入不能为空，请检查!!"));

		goto GameOverDO1;

		return TRUE;
	}


	for(size_t i=0;i<vec_float.size();i++)
	{

		if((int)vec_float[i]==8)
		{
			str_temp=_T("∞");
		}
		else
		{
			str_temp.Format(_T("%.4f"),vec_float[i]);
		}
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//将填入数读取



	vec_float=pMain->CStringtoFloat(pMain->m_edit13);//
	for(size_t i=0;i<vec_float.size();i++)
	{

		if((int)vec_float[i]==8)
		{
			str_temp=_T("∞");
		}
		else
		{
			str_temp.Format(_T("%.4f"),vec_float[i]);
		}
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//将填入数读取
	
	vec_float=pMain->CStringtoFloat(pMain->m_edit14);//

	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.4f"),vec_float[i]);

		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//将填入数读取


	for(size_t i=0;i<23;i++)
	{
		vec_str.push_back(_T("√"));
	}


	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);



	vec_Pos.push_back(_T("D3"));
	vec_Pos.push_back(_T("D4"));
	vec_Pos.push_back(_T("D5"));
	vec_Pos.push_back(_T("D10"));
	vec_Pos.push_back(_T("D11"));
	vec_Pos.push_back(_T("D12"));
	vec_Pos.push_back(_T("D13"));

	vec_Pos.push_back(_T("E19"));
	vec_Pos.push_back(_T("E20"));
	vec_Pos.push_back(_T("E21"));
	vec_Pos.push_back(_T("E22"));
	vec_Pos.push_back(_T("E23"));
	vec_Pos.push_back(_T("E24"));
	vec_Pos.push_back(_T("E25"));
	vec_Pos.push_back(_T("E26"));
	vec_Pos.push_back(_T("E27"));
	vec_Pos.push_back(_T("E28"));
	vec_Pos.push_back(_T("E29"));
	vec_Pos.push_back(_T("E30"));
	vec_Pos.push_back(_T("E31"));
	vec_Pos.push_back(_T("E32"));
	vec_Pos.push_back(_T("E33"));
	vec_Pos.push_back(_T("E34"));	


	vec_Pos.push_back(_T("F3"));
	vec_Pos.push_back(_T("F4"));
	vec_Pos.push_back(_T("F5"));
	vec_Pos.push_back(_T("F10"));
	vec_Pos.push_back(_T("F11"));
	vec_Pos.push_back(_T("F12"));
	vec_Pos.push_back(_T("F13"));

	vec_Pos.push_back(_T("G19"));
	vec_Pos.push_back(_T("G20"));
	vec_Pos.push_back(_T("G21"));
	vec_Pos.push_back(_T("G22"));
	vec_Pos.push_back(_T("G23"));
	vec_Pos.push_back(_T("G24"));
	vec_Pos.push_back(_T("G25"));
	vec_Pos.push_back(_T("G26"));
	vec_Pos.push_back(_T("G27"));
	vec_Pos.push_back(_T("G28"));
	vec_Pos.push_back(_T("G29"));
	vec_Pos.push_back(_T("G30"));
	vec_Pos.push_back(_T("G31"));
	vec_Pos.push_back(_T("G32"));
	vec_Pos.push_back(_T("G33"));
	vec_Pos.push_back(_T("G34"));	

	vec_Pos.push_back(_T("B6"));
	vec_Pos.push_back(_T("D6"));
	vec_Pos.push_back(_T("G6"));	

	vec_Pos.push_back(_T("B14"));
	vec_Pos.push_back(_T("D14"));
	vec_Pos.push_back(_T("G14"));

	vec_Pos.push_back(_T("B35"));
	vec_Pos.push_back(_T("D35"));
	vec_Pos.push_back(_T("G35"));

	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("输入数据数量有误，请检查!"));

		goto GameOverDO1;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);

GameOverDO1:
	    pMain->books.Close(); 
        pMain->app.Quit();  			// 退出
	    //释放对象  
	    pMain->range.ReleaseDispatch();
	    pMain->sheet.ReleaseDispatch();
	    pMain->sheets.ReleaseDispatch();
	    pMain->book.ReleaseDispatch();
	    pMain->books.ReleaseDispatch();
	    pMain->app.ReleaseDispatch();
		CoUninitialize(); 

	    pMain->m_click=TRUE;
	return TRUE;

}

HBRUSH DO::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);



	// TODO:  在此更改 DC 的任何特性

	// TODO:  如果默认的不是所需画笔，则返回另一个画笔
	 if(pWnd->GetDlgCtrlID()== IDC_EDIT12||
		 pWnd->GetDlgCtrlID()== IDC_EDIT13||
		 pWnd->GetDlgCtrlID()== IDC_EDIT14||
		 pWnd->GetDlgCtrlID()== IDC_EDIT15||pWnd->GetDlgCtrlID()== IDC_EDIT16||
		 pWnd->GetDlgCtrlID()== IDC_EDIT17||
		 pWnd->GetDlgCtrlID()== IDC_EDIT18)
	 {
		 pDC->SetTextColor(RGB(0,0,0)); //文字颜色  
		 pDC->SetBkColor(RGB(53,203,60));
		 pDC->SetBkMode(TRANSPARENT);
		 return (HBRUSH) m_Brush.GetSafeHandle();
	 }




	return hbr;
}


void DO::OnBnClickedButton10()
{
	// TODO: 在此添加控件通知处理程序代码

	if(m_click==TRUE)
	{
		m_click=FALSE;
	    AfxBeginThread((AFX_THREADPROC)DO::Button3Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("文件未关闭，不允许操作，请先关闭文件！"));
	}


}

DWORD DO::Button3Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	DO *pMain=(DO*) lpParam;	

	pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);

	m_inputopname=pMain->m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!pMain->app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//得到Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//得到Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)4));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_edit15);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("输入不能为空，请检查!!"));

		goto GameOverDO2;

		return TRUE;
	}


	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.4f"),vec_float[i]);

		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//将填入数读取

	for(size_t i=0;i<16;i++)
	{
		vec_str.push_back(_T("√"));
	}


	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);





	vec_Pos.push_back(_T("F3"));
	vec_Pos.push_back(_T("F4"));
	vec_Pos.push_back(_T("F5"));
	vec_Pos.push_back(_T("F6"));
	vec_Pos.push_back(_T("F7"));
	vec_Pos.push_back(_T("F8"));
	vec_Pos.push_back(_T("F9"));
	vec_Pos.push_back(_T("F10"));
	vec_Pos.push_back(_T("F11"));
	vec_Pos.push_back(_T("F12"));
	vec_Pos.push_back(_T("F13"));
	vec_Pos.push_back(_T("F14"));
	vec_Pos.push_back(_T("F15"));
	vec_Pos.push_back(_T("F16"));
	vec_Pos.push_back(_T("F17"));
	vec_Pos.push_back(_T("F18"));

	vec_Pos.push_back(_T("I3"));
	vec_Pos.push_back(_T("I4"));
	vec_Pos.push_back(_T("I5"));
	vec_Pos.push_back(_T("I6"));
	vec_Pos.push_back(_T("I7"));
	vec_Pos.push_back(_T("I8"));
	vec_Pos.push_back(_T("I9"));
	vec_Pos.push_back(_T("I10"));
	vec_Pos.push_back(_T("I11"));
	vec_Pos.push_back(_T("I12"));
	vec_Pos.push_back(_T("I13"));
	vec_Pos.push_back(_T("I14"));
	vec_Pos.push_back(_T("I15"));
	vec_Pos.push_back(_T("I16"));
	vec_Pos.push_back(_T("I17"));
	vec_Pos.push_back(_T("I18"));

	vec_Pos.push_back(_T("B19"));
	vec_Pos.push_back(_T("D19"));
	vec_Pos.push_back(_T("G19"));


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("输入数据数量有误，请检查!"));

		goto GameOverDO2;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);

GameOverDO2:
	    pMain->books.Close(); 
        pMain->app.Quit();  			// 退出
	    //释放对象  
	    pMain->range.ReleaseDispatch();
	    pMain->sheet.ReleaseDispatch();
	    pMain->sheets.ReleaseDispatch();
	    pMain->book.ReleaseDispatch();
	    pMain->books.ReleaseDispatch();
	    pMain->app.ReleaseDispatch();
		CoUninitialize(); 

	    pMain->m_click=TRUE;
	return TRUE;

}




void DO::OnBnClickedButton11()
{
	// TODO: 在此添加控件通知处理程序代码

	if(m_click==TRUE)
	{
		m_click=FALSE;
	    AfxBeginThread((AFX_THREADPROC)DO::Button11Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("文件未关闭，不允许操作，请先关闭文件！"));
	}

}


void DO::OnBnClickedButton12()
{
	// TODO: 在此添加控件通知处理程序代码

	if(m_click==TRUE)
	{
		m_click=FALSE;
	    AfxBeginThread((AFX_THREADPROC)DO::Button12Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("文件未关闭，不允许操作，请先关闭文件！"));
	}
}


DWORD DO::Button11Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	DO *pMain=(DO*) lpParam;	

	pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);

	m_inputopname=pMain->m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!pMain->app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//得到Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//得到Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)5));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_edit16);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("输入不能为空，请检查!!"));

		goto GameOverDO3;

		return TRUE;
	}


	for(size_t i=0;i<vec_float.size();i++)
	{


		if((int)vec_float[i]==8)
		{
			str_temp=_T("∞");
		}
		else
		{
			str_temp.Format(_T("%.4f"),vec_float[i]);
		}
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//将填入数读取

	for(size_t i=0;i<16;i++)
	{
		vec_str.push_back(_T("√"));
	}


	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);





	vec_Pos.push_back(_T("F3"));
	vec_Pos.push_back(_T("F4"));
	vec_Pos.push_back(_T("F5"));
	vec_Pos.push_back(_T("F6"));
	vec_Pos.push_back(_T("F7"));
	vec_Pos.push_back(_T("F8"));
	vec_Pos.push_back(_T("F9"));
	vec_Pos.push_back(_T("F10"));
	vec_Pos.push_back(_T("F11"));
	vec_Pos.push_back(_T("F12"));
	vec_Pos.push_back(_T("F13"));
	vec_Pos.push_back(_T("F14"));
	vec_Pos.push_back(_T("F15"));
	vec_Pos.push_back(_T("F16"));
	vec_Pos.push_back(_T("F17"));
	vec_Pos.push_back(_T("F18"));

	vec_Pos.push_back(_T("H3"));
	vec_Pos.push_back(_T("H4"));
	vec_Pos.push_back(_T("H5"));
	vec_Pos.push_back(_T("H6"));
	vec_Pos.push_back(_T("H7"));
	vec_Pos.push_back(_T("H8"));
	vec_Pos.push_back(_T("H9"));
	vec_Pos.push_back(_T("H10"));
	vec_Pos.push_back(_T("H11"));
	vec_Pos.push_back(_T("H12"));
	vec_Pos.push_back(_T("H13"));
	vec_Pos.push_back(_T("H14"));
	vec_Pos.push_back(_T("H15"));
	vec_Pos.push_back(_T("H16"));
	vec_Pos.push_back(_T("H17"));
	vec_Pos.push_back(_T("H18"));

	vec_Pos.push_back(_T("B19"));
	vec_Pos.push_back(_T("D19"));
	vec_Pos.push_back(_T("G19"));


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("输入数据数量有误，请检查!"));

		goto GameOverDO3;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);

GameOverDO3:
	    pMain->books.Close(); 
        pMain->app.Quit();  			// 退出
	    //释放对象  
	    pMain->range.ReleaseDispatch();
	    pMain->sheet.ReleaseDispatch();
	    pMain->sheets.ReleaseDispatch();
	    pMain->book.ReleaseDispatch();
	    pMain->books.ReleaseDispatch();
	    pMain->app.ReleaseDispatch();
		CoUninitialize(); 

	    pMain->m_click=TRUE;
	return TRUE;

}


DWORD DO::Button12Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	DO *pMain=(DO*) lpParam;	

	pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);

	m_inputopname=pMain->m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!pMain->app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//得到Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//得到Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)6));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_edit17);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("输入不能为空，请检查!!"));

		goto GameOverDO4;

		return TRUE;
	}


	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.4f"),vec_float[i]);

		vec_str.push_back(str_temp);
	    str_temp=_T("");
	 }//将填入数读取

	vec_float=pMain->CStringtoFloat(pMain->m_edit18);//

	if(vec_float.empty())
	{
		AfxMessageBox(_T("输入不能为空，请检查!!"));

		goto GameOverDO4;

		return TRUE;
	}

	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.4f"),vec_float[i]);

		vec_str.push_back(str_temp);
	    str_temp=_T("");
	 }//将填入数读取

	for(size_t i=0;i<8;i++)
	{
		vec_str.push_back(_T("×"));
		vec_str.push_back(_T("√"));
		vec_str.push_back(_T("√"));
		vec_str.push_back(_T("√"));
	}

	for(size_t i=0;i<35;i++)
	{
		vec_str.push_back(_T("√"));
	}

	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);

	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);




	vec_Pos.push_back(_T("F3"));
	vec_Pos.push_back(_T("F4"));
	vec_Pos.push_back(_T("F5"));



	vec_Pos.push_back(_T("E10"));
	vec_Pos.push_back(_T("E11"));
	vec_Pos.push_back(_T("E12"));
	vec_Pos.push_back(_T("E13"));
	vec_Pos.push_back(_T("E14"));
	vec_Pos.push_back(_T("E15"));
	vec_Pos.push_back(_T("E16"));
	vec_Pos.push_back(_T("E17"));
	vec_Pos.push_back(_T("E18"));

	vec_Pos.push_back(_T("E19"));
	vec_Pos.push_back(_T("E20"));
	vec_Pos.push_back(_T("E21"));
	vec_Pos.push_back(_T("E22"));
	vec_Pos.push_back(_T("E23"));
	vec_Pos.push_back(_T("E24"));
	vec_Pos.push_back(_T("E25"));
	vec_Pos.push_back(_T("E26"));
	vec_Pos.push_back(_T("E27"));

	vec_Pos.push_back(_T("E28"));
	vec_Pos.push_back(_T("E29"));
	vec_Pos.push_back(_T("E30"));
	vec_Pos.push_back(_T("E31"));
	vec_Pos.push_back(_T("E32"));
	vec_Pos.push_back(_T("E33"));
	vec_Pos.push_back(_T("E34"));
	vec_Pos.push_back(_T("E35"));
	vec_Pos.push_back(_T("E36"));

	vec_Pos.push_back(_T("E37"));
	vec_Pos.push_back(_T("E38"));
	vec_Pos.push_back(_T("E39"));
	vec_Pos.push_back(_T("E40"));
	vec_Pos.push_back(_T("E41"));



	vec_Pos.push_back(_T("D10"));
	vec_Pos.push_back(_T("D11"));
	vec_Pos.push_back(_T("D12"));
	vec_Pos.push_back(_T("D13"));
	vec_Pos.push_back(_T("D14"));
	vec_Pos.push_back(_T("D15"));
	vec_Pos.push_back(_T("D16"));
	vec_Pos.push_back(_T("D17"));
	vec_Pos.push_back(_T("D18"));

	vec_Pos.push_back(_T("D19"));
	vec_Pos.push_back(_T("D20"));
	vec_Pos.push_back(_T("D21"));
	vec_Pos.push_back(_T("D22"));
	vec_Pos.push_back(_T("D23"));
	vec_Pos.push_back(_T("D24"));
	vec_Pos.push_back(_T("D25"));
	vec_Pos.push_back(_T("D26"));
	vec_Pos.push_back(_T("D27"));

	vec_Pos.push_back(_T("D28"));
	vec_Pos.push_back(_T("D29"));
	vec_Pos.push_back(_T("D30"));
	vec_Pos.push_back(_T("D31"));
	vec_Pos.push_back(_T("D32"));
	vec_Pos.push_back(_T("D33"));
	vec_Pos.push_back(_T("D34"));
	vec_Pos.push_back(_T("D35"));
	vec_Pos.push_back(_T("D36"));

	vec_Pos.push_back(_T("D37"));
	vec_Pos.push_back(_T("D38"));
	vec_Pos.push_back(_T("D39"));
	vec_Pos.push_back(_T("D40"));
	vec_Pos.push_back(_T("D41"));


	vec_Pos.push_back(_T("G3"));
	vec_Pos.push_back(_T("G4"));
	vec_Pos.push_back(_T("G5"));

	vec_Pos.push_back(_T("G10"));
	vec_Pos.push_back(_T("G11"));
	vec_Pos.push_back(_T("G12"));
	vec_Pos.push_back(_T("G13"));
	vec_Pos.push_back(_T("G14"));
	vec_Pos.push_back(_T("G15"));
	vec_Pos.push_back(_T("G16"));
	vec_Pos.push_back(_T("G17"));
	vec_Pos.push_back(_T("G18"));

	vec_Pos.push_back(_T("G19"));
	vec_Pos.push_back(_T("G20"));
	vec_Pos.push_back(_T("G21"));
	vec_Pos.push_back(_T("G22"));
	vec_Pos.push_back(_T("G23"));
	vec_Pos.push_back(_T("G24"));
	vec_Pos.push_back(_T("G25"));
	vec_Pos.push_back(_T("G26"));
	vec_Pos.push_back(_T("G27"));

	vec_Pos.push_back(_T("G28"));
	vec_Pos.push_back(_T("G29"));
	vec_Pos.push_back(_T("G30"));
	vec_Pos.push_back(_T("G31"));
	vec_Pos.push_back(_T("G32"));
	vec_Pos.push_back(_T("G33"));
	vec_Pos.push_back(_T("G34"));
	vec_Pos.push_back(_T("G35"));
	vec_Pos.push_back(_T("G36"));

	vec_Pos.push_back(_T("G37"));
	vec_Pos.push_back(_T("G38"));
	vec_Pos.push_back(_T("G39"));
	vec_Pos.push_back(_T("G40"));
	vec_Pos.push_back(_T("G41"));

	vec_Pos.push_back(_T("B6"));
	vec_Pos.push_back(_T("D6"));
	vec_Pos.push_back(_T("G6"));


	vec_Pos.push_back(_T("B42"));
	vec_Pos.push_back(_T("D42"));
	vec_Pos.push_back(_T("G42"));


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("输入数据数量有误，请检查!"));

		goto GameOverDO4;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)7));

	vec_str.clear();
	vec_Pos.clear();

	for(size_t i=0;i<8;i++)
	{
		vec_str.push_back(_T("是"));
	}

	for(size_t i=0;i<8;i++)
	{
		vec_str.push_back(_T("√"));
	}

	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);

	vec_Pos.push_back(_T("E3"));
	vec_Pos.push_back(_T("E4"));
	vec_Pos.push_back(_T("E5"));
	vec_Pos.push_back(_T("E6"));
	vec_Pos.push_back(_T("E7"));
	vec_Pos.push_back(_T("E8"));
	vec_Pos.push_back(_T("E9"));
	vec_Pos.push_back(_T("E10"));


	vec_Pos.push_back(_T("G3"));
	vec_Pos.push_back(_T("G4"));
	vec_Pos.push_back(_T("G5"));
	vec_Pos.push_back(_T("G6"));
	vec_Pos.push_back(_T("G7"));
	vec_Pos.push_back(_T("G8"));
	vec_Pos.push_back(_T("G9"));
	vec_Pos.push_back(_T("G10"));

	
	vec_Pos.push_back(_T("B11"));
	vec_Pos.push_back(_T("D11"));
	vec_Pos.push_back(_T("G11"));

	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("输入数据数量有误，请检查!"));

		goto GameOverDO4;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}




	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);

GameOverDO4:
	    pMain->books.Close(); 
        pMain->app.Quit();  			// 退出
	    //释放对象  
	    pMain->range.ReleaseDispatch();
	    pMain->sheet.ReleaseDispatch();
	    pMain->sheets.ReleaseDispatch();
	    pMain->book.ReleaseDispatch();
	    pMain->books.ReleaseDispatch();
	    pMain->app.ReleaseDispatch();
		CoUninitialize(); 

	    pMain->m_click=TRUE;
	return TRUE;

}


void DO::OnBnClickedCancel()
{
	// TODO: 在此添加控件通知处理程序代码

	m_closeornot=TRUE;

	if(AfxMessageBox(_T("您确定要退出当前DO板填表工作，劝君三思而行!"),MB_YESNO|MB_ICONEXCLAMATION)==IDYES)
	{
		CDialogEx::OnCancel();
	}

	
}

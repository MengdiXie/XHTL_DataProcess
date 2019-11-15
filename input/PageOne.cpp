// PageOne.cpp : 实现文件
//

#include "stdafx.h"
#include "input.h"
#include "PageOne.h"
#include "afxdialogex.h"


extern CString exfilename;//外部变量
extern CString m_inputopname;

// PageOne 对话框

IMPLEMENT_DYNAMIC(PageOne, CDialogEx)

PageOne::PageOne(CWnd* pParent /*=NULL*/)
	: CDialogEx(PageOne::IDD, pParent)
{

}

PageOne::~PageOne()
{
}

void PageOne::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Check(pDX, IDC_dui1, m_dui1);
	DDX_Check(pDX, IDC_dui2, m_dui2);
	DDX_Check(pDX, IDC_dui3, m_dui3);
	DDX_Check(pDX, IDC_dui4, m_dui4);
	DDX_Check(pDX, IDC_dui5, m_dui5);
	DDX_Check(pDX, IDC_dui6, m_dui6);
	DDX_Check(pDX, IDC_dui7, m_dui7);
	DDX_Check(pDX, IDC_dui8, m_dui8);
	DDX_Check(pDX, IDC_dui9, m_dui9);
	DDX_Check(pDX, IDC_dui10, m_dui10);
	DDX_Check(pDX, IDC_dui11, m_dui11);
	DDX_Check(pDX, IDC_dui12, m_dui12);
	DDX_Check(pDX, IDC_cuo1, m_cuo1);
	DDX_Check(pDX, IDC_cuo2, m_cuo2);
	DDX_Check(pDX, IDC_cuo3, m_cuo3);
	DDX_Check(pDX, IDC_cuo4, m_cuo4);
	DDX_Check(pDX, IDC_cuo5, m_cuo5);
	DDX_Check(pDX, IDC_cuo6, m_cuo6);
	DDX_Check(pDX, IDC_cuo7, m_cuo7);
	DDX_Check(pDX, IDC_cuo8, m_cuo8);
	DDX_Check(pDX, IDC_cuo9, m_cuo9);
	DDX_Check(pDX, IDC_cuo10, m_cuo10);
	DDX_Check(pDX, IDC_cuo11, m_cuo11);
	DDX_Check(pDX, IDC_cuo12, m_cuo12);

	DDX_Text(pDX, IDC_INPUT, m_inputstr);
	DDX_Text(pDX, IDC_opname, m_opname);
	DDX_Text(pDX, IDC_INPUT2, m_inputstr2);
	DDX_Text(pDX, IDC_INPUT3, m_inputstr3);
	DDX_Text(pDX, IDC_INPUT4, m_inputstr4);
	DDX_Text(pDX, IDC_INPUT5, m_inputstr5);
	DDX_Text(pDX, IDC_INPUT6, m_inputstr6);
	DDX_Text(pDX, IDC_INPUT7, m_inputstr7);
	DDX_Text(pDX, IDC_INPUT8, m_inputstr8);


	DDX_Text(pDX, IDC_EDIT1, m_boardnum);


	DDX_Text(pDX, IDC_EDIT2, m_tempture);
	DDX_Text(pDX, IDC_EDIT3, m_humidity);

	DDX_Text(pDX, IDC_EDIT4, m_multimeter);
	DDX_Text(pDX, IDC_EDIT5, m_constant_volt_source);

	DDX_Text(pDX, IDC_EDIT6, m_signal_source);
	DDX_Text(pDX, IDC_EDIT9, m_测试工装);

}


BEGIN_MESSAGE_MAP(PageOne, CDialogEx)
	ON_BN_CLICKED(IDOK, &PageOne::OnBnClickedOk)
	ON_BN_CLICKED(IDC_two, &PageOne::OnBnClickedtwo)
	ON_BN_CLICKED(IDC_three, &PageOne::OnBnClickedthree)
	ON_BN_CLICKED(IDC_stop, &PageOne::OnBnClickedstop)
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_BUTTON1, &PageOne::OnBnClickedButton1)
	ON_BN_CLICKED(IDCANCEL, &PageOne::OnBnClickedCancel)
END_MESSAGE_MAP()


// PageOne 消息处理程序


BOOL PageOne::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化


	//sheet = sheets.get_Item(COleVariant((short)1));
	m_dui1=true;
	m_dui2=true;
	m_dui3=true;
	m_dui4=true;
	m_dui5=true;
	m_dui6=true;
	m_dui7=true;
	m_dui8=true;
	m_dui9=true;
	m_dui10=true;
	m_dui11=true;
	m_dui12=true;

	my_Font.CreatePointFont(140, L"Times NewRoman");//创建字体和大小

	m_Brush.CreateSolidBrush(RGB(53,203,60));

	 GetDlgItem(IDC_INPUT)->SetFont(&my_Font);
	 GetDlgItem(IDC_INPUT2)->SetFont(&my_Font);
	 GetDlgItem(IDC_INPUT3)->SetFont(&my_Font);
	 GetDlgItem(IDC_INPUT4)->SetFont(&my_Font);
	 GetDlgItem(IDC_INPUT5)->SetFont(&my_Font);
	 GetDlgItem(IDC_INPUT6)->SetFont(&my_Font);
	 GetDlgItem(IDC_INPUT7)->SetFont(&my_Font);
	 GetDlgItem(IDC_INPUT8)->SetFont(&my_Font);



	m_boardnum=_T("U-I_201901001");
	m_opname=_T("路人甲");
	m_tempture=_T("25℃");
	m_humidity=_T("40%");

	 m_multimeter=_T("20190620-20200619");
	 m_constant_volt_source=_T("20190620-20200619");
	 m_signal_source=_T("20190620-20200619");
	 m_测试工装=_T("20190620-20200619");
	 m_inputstr2=_T("8,8,8,8,8,8,8,8");
	 m_inputstr3=_T("2.7,8,8,8,8,8,8,2.7,8,8,8,8,8,2.7,8,8,8,8,8,8,144,8,144,142");


	UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void PageOne::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );

	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return;
	}
	books.AttachDispatch(app.get_Workbooks());
	lpDisp = books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);

	
	
	//得到Workbook
	book.AttachDispatch(lpDisp);
	//得到Worksheets
	sheets.AttachDispatch(book.get_Worksheets());

	UpdateData(TRUE);
	if(m_dui1==true)
	{
		m_cuo2=false;
	}
#if 1
	sheet = sheets.get_Item(COleVariant((short)1));
	//sheet.AttachDispatch(lpDisp);
	//选择需要写入的单元格
	lpDisp = sheet.get_Range(COleVariant(_T("F3")), COleVariant(_T("F3")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));
	lpDisp = sheet.get_Range(COleVariant(_T("F4")), COleVariant(_T("F4")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));

	lpDisp = sheet.get_Range(COleVariant(_T("F5")), COleVariant(_T("F5")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));

	lpDisp = sheet.get_Range(COleVariant(_T("F6")), COleVariant(_T("F6")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));



	
	lpDisp = sheet.get_Range(COleVariant(_T("G12")), COleVariant(_T("G12")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));
	lpDisp = sheet.get_Range(COleVariant(_T("G13")), COleVariant(_T("G13")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));
	lpDisp = sheet.get_Range(COleVariant(_T("G14")), COleVariant(_T("G14")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));
	lpDisp = sheet.get_Range(COleVariant(_T("G15")), COleVariant(_T("G15")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));
	lpDisp = sheet.get_Range(COleVariant(_T("G16")), COleVariant(_T("G16")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));
	lpDisp = sheet.get_Range(COleVariant(_T("G17")), COleVariant(_T("G17")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));
	lpDisp = sheet.get_Range(COleVariant(_T("G18")), COleVariant(_T("G18")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));
	lpDisp = sheet.get_Range(COleVariant(_T("G19")), COleVariant(_T("G19")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));

	vec_float=CStringtoFloat(m_inputstr);


	CString tempstr;
	for(size_t i=0;i<vec_float.size();++i)
	{
		tempstr.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(tempstr);
		tempstr=_T("");
	}
	std::vector<CString>  Pos;
	Pos.push_back(_T("F12"));Pos.push_back(_T("F13"));
	Pos.push_back(_T("F14"));Pos.push_back(_T("F15"));
	Pos.push_back(_T("F16"));Pos.push_back(_T("F17"));
	Pos.push_back(_T("F18"));Pos.push_back(_T("F19"));
	for(size_t i=0;i<Pos.size();i++)
	{
		lpDisp = sheet.get_Range(COleVariant(Pos[i]), COleVariant(Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}


	lpDisp = sheet.get_Range(COleVariant(_T("C7")), COleVariant(_T("C7")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(m_opname));
	lpDisp = sheet.get_Range(COleVariant(_T("C20")), COleVariant(_T("C20")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(m_opname));


	lpDisp = sheet.get_Range(COleVariant(_T("G7")), COleVariant(_T("G7")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("是否合格：合格√  不合格")));//打印合格
	lpDisp = sheet.get_Range(COleVariant(_T("G20")), COleVariant(_T("G20")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("是否合格：合格√  不合格")));//打印合格


	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//打印填表日期
	lpDisp = sheet.get_Range(COleVariant(_T("F7")), COleVariant(_T("F7")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(strtime));
	lpDisp = sheet.get_Range(COleVariant(_T("F20")), COleVariant(_T("F20")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(strtime));




	m_inputopname=m_opname;



	app.put_Visible(TRUE);
	book.Save();
	book.put_Saved(TRUE);



	books.Close(); 
    app.Quit();  			// 退出
	//释放对象  
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.ReleaseDispatch();

	CoUninitialize(); 
#if 0

	if(!app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return;
	}
	books.AttachDispatch(app.get_Workbooks());
	lpDisp = books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//得到Workbook
	book.AttachDispatch(lpDisp);
	//得到Worksheets
	sheets.AttachDispatch(book.get_Worksheets());

     sheet = sheets.get_Item(COleVariant((short)1));
	//sheet.AttachDispatch(lpDisp);
	//选择需要写入的单元格
	lpDisp = sheet.get_Range(COleVariant(_T("G12")), COleVariant(_T("G12")));
	//将数据链接到单元格//将数据写入对应的单元格
	range.AttachDispatch(lpDisp);
	range.put_Value(vtMissing, COleVariant(_T("√")));


	
	books.Close(); 
    app.Quit();  			// 退出
	//释放对象  
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.ReleaseDispatch();

#endif
#endif
	
}


std::vector<float> PageOne::CStringtoFloat(CString m_s)
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

void PageOne::OnBnClickedtwo()
{
	// TODO: 在此添加控件通知处理程序代码

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );

	UpdateData(TRUE);
	m_inputopname=m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return;
	}
	books.AttachDispatch(app.get_Workbooks());
	lpDisp = books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//得到Workbook
	book.AttachDispatch(lpDisp);
	//得到Worksheets
	sheets.AttachDispatch(book.get_Worksheets());

	sheet = sheets.get_Item(COleVariant((short)2));
	CString str_temp; 
	vec_float=CStringtoFloat(m_inputstr2);//



	
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

	vec_Pos.push_back(_T("H3"));
	vec_Pos.push_back(_T("H4"));
	vec_Pos.push_back(_T("H5"));
	vec_Pos.push_back(_T("H6"));
	vec_Pos.push_back(_T("H7"));
	vec_Pos.push_back(_T("H8"));
	vec_Pos.push_back(_T("H9"));
	vec_Pos.push_back(_T("H10"));
	if(vec_Pos.size()!=vec_float.size())
	{
		MessageBox(_T("输入数据数量与实际数量不符，请重新输入!"));
        books.Close(); 
        app.Quit();  			// 退出
	    //释放对象  
	    range.ReleaseDispatch();
	    sheet.ReleaseDispatch();
	    sheets.ReleaseDispatch();
	    book.ReleaseDispatch();
	    books.ReleaseDispatch();
	    app.ReleaseDispatch();


		return;
	}
	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}
	vec_Pos.clear();
	vec_Pos.push_back(_T("I3"));
	vec_Pos.push_back(_T("I4"));
	vec_Pos.push_back(_T("I5"));
	vec_Pos.push_back(_T("I6"));
	vec_Pos.push_back(_T("I7"));
	vec_Pos.push_back(_T("I8"));
	vec_Pos.push_back(_T("I9"));
	vec_Pos.push_back(_T("I10"));
	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(_T("√")));
	}

	vec_Pos.clear();
	vec_str.clear();

	vec_Pos.push_back(_T("C11"));
	vec_Pos.push_back(_T("F11"));
	vec_Pos.push_back(_T("I11"));
	vec_str.push_back(m_opname);

	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//打印填表日期

	vec_str.push_back(strtime);
	vec_str.push_back(_T("合格√ 不合格"));


	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}


	app.put_Visible(TRUE);
	book.Save();
	book.put_Saved(TRUE);



    books.Close(); 
    app.Quit();  			// 退出
	//释放对象  
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.ReleaseDispatch();


   CoUninitialize(); 


}


void PageOne::OnBnClickedthree()
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );
	int count=0;
	UpdateData(TRUE);
	m_inputopname=m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return;
	}
	books.AttachDispatch(app.get_Workbooks());
	lpDisp = books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//得到Workbook
	book.AttachDispatch(lpDisp);
	//得到Worksheets
	sheets.AttachDispatch(book.get_Worksheets());

	sheet = sheets.get_Item(COleVariant((short)4));
	CString str_temp; 
	vec_float=CStringtoFloat(m_inputstr4);//



	count+=vec_float.size();
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
		str_temp=_T("");
	}

	vec_float=CStringtoFloat(m_inputstr5);//


	count+=vec_float.size();
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%f"),vec_float[i]);
		vec_str.push_back(str_temp);
		str_temp=_T("");
	}
#if 1
	vec_float=CStringtoFloat(m_inputstr6);//


	count+=vec_float.size();
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
		str_temp=_T("");
	}
	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//打印填表日期

	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("合格√ 不合格"));
	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("合格√ 不合格"));
	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("合格√ 不合格"));


	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));

#endif



	vec_Pos.push_back(_T("I3"));
	vec_Pos.push_back(_T("I4"));
	vec_Pos.push_back(_T("I5"));
	vec_Pos.push_back(_T("I6"));

	vec_Pos.push_back(_T("I13"));
	vec_Pos.push_back(_T("I14"));
	vec_Pos.push_back(_T("I15"));
	vec_Pos.push_back(_T("I16"));
	vec_Pos.push_back(_T("I17"));
	vec_Pos.push_back(_T("I18"));
	vec_Pos.push_back(_T("I19"));
	vec_Pos.push_back(_T("I20"));
	vec_Pos.push_back(_T("I21"));
#if 1
	vec_Pos.push_back(_T("I28"));
	vec_Pos.push_back(_T("I29"));
	vec_Pos.push_back(_T("I30"));

	vec_Pos.push_back(_T("C7"));
	vec_Pos.push_back(_T("F7"));
	vec_Pos.push_back(_T("I7"));
	vec_Pos.push_back(_T("C22"));
	vec_Pos.push_back(_T("F22"));
	vec_Pos.push_back(_T("I22"));
	vec_Pos.push_back(_T("C31"));
	vec_Pos.push_back(_T("F31"));
	vec_Pos.push_back(_T("I31"));
#endif


	vec_Pos.push_back(_T("J3"));
	vec_Pos.push_back(_T("J4"));
	vec_Pos.push_back(_T("J5"));
	vec_Pos.push_back(_T("J6"));

	vec_Pos.push_back(_T("J13"));
	vec_Pos.push_back(_T("J14"));
	vec_Pos.push_back(_T("J15"));
	vec_Pos.push_back(_T("J16"));
	vec_Pos.push_back(_T("J17"));
	vec_Pos.push_back(_T("J18"));
	vec_Pos.push_back(_T("J19"));
	vec_Pos.push_back(_T("J20"));
	vec_Pos.push_back(_T("J21"));
#if 1
	vec_Pos.push_back(_T("J28"));
	vec_Pos.push_back(_T("J29"));
	vec_Pos.push_back(_T("J30"));
#endif


	if(vec_Pos.size()!=vec_str.size())
	{
		MessageBox(_T("输入数据数量与实际数量不符，请重新输入!"));
        books.Close(); 
        app.Quit();  			// 退出
	    //释放对象  
	    range.ReleaseDispatch();
	    sheet.ReleaseDispatch();
	    sheets.ReleaseDispatch();
	    book.ReleaseDispatch();
	    books.ReleaseDispatch();
	    app.ReleaseDispatch();


		return;
	}




	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
		
	}
	
	vec_str.clear();
	vec_Pos.clear();
	sheet = sheets.get_Item(COleVariant((short)5));

	vec_float=CStringtoFloat(m_inputstr7);//
	count+=vec_float.size();
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
		str_temp=_T("");
	}

	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));

	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("合格√ 不合格"));
	
	vec_Pos.push_back(_T("F4"));
	vec_Pos.push_back(_T("F5"));
	vec_Pos.push_back(_T("K4"));
	vec_Pos.push_back(_T("K5"));

	vec_Pos.push_back(_T("L4"));
	vec_Pos.push_back(_T("L5"));



	vec_Pos.push_back(_T("D6"));
	vec_Pos.push_back(_T("H6"));

	vec_Pos.push_back(_T("K6"));



	if(vec_Pos.size()!=vec_str.size())
	{
		MessageBox(_T("输入数据数量与实际数量不符，请重新输入!"));
        books.Close(); 
        app.Quit();  			// 退出
	    //释放对象  
	    range.ReleaseDispatch();
	    sheet.ReleaseDispatch();
	    sheets.ReleaseDispatch();
	    book.ReleaseDispatch();
	    books.ReleaseDispatch();
	    app.ReleaseDispatch();


		return;
	}




	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
		
	}




	vec_str.clear();
	vec_Pos.clear();
	sheet = sheets.get_Item(COleVariant((short)6));

	vec_float=CStringtoFloat(m_inputstr8);//
	count+=vec_float.size();
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
		str_temp=_T("");
	}

	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("合格√ 不合格"));


	vec_Pos.push_back(_T("K4"));
	vec_Pos.push_back(_T("K5"));
	vec_Pos.push_back(_T("K6"));
	vec_Pos.push_back(_T("K7"));
	vec_Pos.push_back(_T("K8"));
	vec_Pos.push_back(_T("K9"));
	vec_Pos.push_back(_T("K10"));
	vec_Pos.push_back(_T("K11"));
	vec_Pos.push_back(_T("K12"));

	vec_Pos.push_back(_T("C13"));
	vec_Pos.push_back(_T("G13"));

	vec_Pos.push_back(_T("K13"));



	if(vec_Pos.size()!=vec_str.size())
	{
		MessageBox(_T("输入数据数量与实际数量不符，请重新输入!"));
        books.Close(); 
        app.Quit();  			// 退出
	    //释放对象  
	    range.ReleaseDispatch();
	    sheet.ReleaseDispatch();
	    sheets.ReleaseDispatch();
	    book.ReleaseDispatch();
	    books.ReleaseDispatch();
	    app.ReleaseDispatch();


		return;
	}




	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
		
	}

	sheet = sheets.get_Item(COleVariant((short)7));
	vec_str.clear();
	vec_Pos.clear();
	vec_str.push_back(_T("√"));
	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("合格√ 不合格"));

	vec_Pos.push_back(_T("G4"));
	vec_Pos.push_back(_T("B5"));
	vec_Pos.push_back(_T("E5"));
	vec_Pos.push_back(_T("H5"));
	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
		
	}
	app.put_Visible(TRUE);
	book.Save();
	book.put_Saved(TRUE);



	books.Close(); 
    app.Quit();  			// 退出
	//释放对象  
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.ReleaseDispatch();
	// TODO: 在此添加控件通知处理程序代码
    CoUninitialize(); 


	CDialogEx::OnOK();

}


void PageOne::OnBnClickedstop()
{
	// TODO: 在此添加控件通知处理程序代码

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );

	UpdateData(TRUE);
	m_inputopname=m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return;
	}
	books.AttachDispatch(app.get_Workbooks());
	lpDisp = books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//得到Workbook
	book.AttachDispatch(lpDisp);
	//得到Worksheets
	sheets.AttachDispatch(book.get_Worksheets());

	sheet = sheets.get_Item(COleVariant((short)3));
	CString str_temp; 
	vec_float=CStringtoFloat(m_inputstr3);//


	
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
	}

	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//打印填表日期

	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("合格√ 不合格"));

	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));

	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));

	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));

	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));

	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));
	vec_str.push_back(_T("√"));

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
	vec_Pos.push_back(_T("H19"));
	vec_Pos.push_back(_T("H20"));


	vec_Pos.push_back(_T("H21"));
	vec_Pos.push_back(_T("H22"));
	vec_Pos.push_back(_T("H23"));
	vec_Pos.push_back(_T("H24"));
	vec_Pos.push_back(_T("H25"));
	vec_Pos.push_back(_T("H26"));


	vec_Pos.push_back(_T("B27"));
	vec_Pos.push_back(_T("E27"));
	vec_Pos.push_back(_T("H27"));

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
	vec_Pos.push_back(_T("I19"));
	vec_Pos.push_back(_T("I20"));


	vec_Pos.push_back(_T("I21"));
	vec_Pos.push_back(_T("I22"));
	vec_Pos.push_back(_T("I23"));
	vec_Pos.push_back(_T("I24"));
	vec_Pos.push_back(_T("I25"));
	vec_Pos.push_back(_T("I26"));

	if(vec_Pos.size()!=vec_str.size())
	{
		MessageBox(_T("输入数据数量与实际数量不符，请重新输入!"));
#if 0
        books.Close(); 
        app.Quit();  			// 退出
	    //释放对象  
	    range.ReleaseDispatch();
	    sheet.ReleaseDispatch();
	    sheets.ReleaseDispatch();
	    book.ReleaseDispatch();
	    books.ReleaseDispatch();
	    app.ReleaseDispatch();

#endif
		goto GameOver;
		return;
	}




	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
		
	}

	app.put_Visible(TRUE);
	book.Save();
	book.put_Saved(TRUE);


GameOver:
	  books.Close(); 
      app.Quit();  			// 退出
	  //释放对象  
	  range.ReleaseDispatch();
	  sheet.ReleaseDispatch();
	  sheets.ReleaseDispatch();
	  book.ReleaseDispatch();
	  books.ReleaseDispatch();
	  app.ReleaseDispatch();

	  CoUninitialize(); 
}


HBRUSH PageOne::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  在此更改 DC 的任何特性
	 if(pWnd->GetDlgCtrlID()== IDC_INPUT||
		 pWnd->GetDlgCtrlID()== IDC_INPUT2||
		 pWnd->GetDlgCtrlID()== IDC_INPUT3||
		 pWnd->GetDlgCtrlID()== IDC_INPUT4||
		 pWnd->GetDlgCtrlID()== IDC_INPUT5||
		 pWnd->GetDlgCtrlID()== IDC_INPUT6||
		 pWnd->GetDlgCtrlID()== IDC_INPUT7||
		 pWnd->GetDlgCtrlID()== IDC_INPUT8)

	 {
		 pDC->SetTextColor(RGB(0,0,0)); //文字颜色  
		 pDC->SetBkColor(RGB(53,203,60));
		 pDC->SetBkMode(TRANSPARENT);
		 return (HBRUSH) m_Brush.GetSafeHandle();
	 }

	// TODO:  如果默认的不是所需画笔，则返回另一个画笔
	return hbr;
}


void PageOne::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );

	std::vector<CString> vec_str;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"无法启动Excel服务器!");
	    return;
	}
	books.AttachDispatch(app.get_Workbooks());
	lpDisp = books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);

	
	
	//得到Workbook
	book.AttachDispatch(lpDisp);
	//得到Worksheets
	sheets.AttachDispatch(book.get_Worksheets());

	

	sheet = sheets.get_Item(COleVariant((short)8));

	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//打印填表日期


	UpdateData(TRUE);
	vec_str.push_back(m_boardnum);
	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("D"));
	vec_str.push_back(_T("合格（  √  ） /  不合格（   ）"));
	vec_str.push_back(m_tempture);
	vec_str.push_back(m_humidity);
	vec_str.push_back(m_multimeter);
	vec_str.push_back(m_constant_volt_source);
	vec_str.push_back(m_signal_source);
	vec_str.push_back(m_测试工装);

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
    vec_Pos.push_back(_T("D16"));


	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //将数据链接到单元格//将数据写入对应的单元格
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	//app.put_Visible(TRUE);
	//book.Save();
	//book.put_Saved(TRUE);


	books.Close(); 
    app.Quit();  			// 退出
	//释放对象  
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.ReleaseDispatch();

	CoUninitialize(); 

}


void PageOne::OnBnClickedCancel()
{
	// TODO: 在此添加控件通知处理程序代码

	if(AfxMessageBox(_T("您确定要退出当前U-I填表工作，请三思而行!"),MB_YESNO|MB_ICONEXCLAMATION)==IDYES)
	{
		CDialogEx::OnCancel();
	}
	
}

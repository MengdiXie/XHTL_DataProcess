// XHTL_ALL.cpp : 实现文件
//

#include "stdafx.h"
#include "input.h"
#include "XHTL_ALL.h"
#include "afxdialogex.h"
#include<algorithm>

// XHTL_ALL 对话框
extern CString exfilename;//外部变量
extern CString m_inputopname;
extern CString m_externinputfilename;
#define WM_UPDATE_MESSAGE2 10008
#define WM_UPDATE_MESSAGE 10007
IMPLEMENT_DYNAMIC(XHTL_ALL, CDialogEx)

XHTL_ALL::XHTL_ALL(CWnd* pParent /*=NULL*/)
	: CDialogEx(XHTL_ALL::IDD, pParent)
{

}

XHTL_ALL::~XHTL_ALL()
{
}

void XHTL_ALL::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Control(pDX,IDC_COMBO1,m_Select1);
	DDX_Check(pDX, IDC_62102, m_checkbox1);
	DDX_Check(pDX, IDC_62103, m_checkbox2);
	DDX_Text(pDX, IDC_EDIT8, m_IDC_EDIT8);
	DDX_Text(pDX, IDC_EDIT9, m_IDC_EDIT9);

	DDX_Control(pDX,IDC_COMBO2,m_Select2);
	DDX_Check(pDX, IDC_62104, m_checkbox3);
	DDX_Check(pDX, IDC_62105, m_checkbox4);
	DDX_Text(pDX, IDC_EDIT14, m_IDC_EDIT14);


	DDX_Text(pDX, IDC_EDIT1, m_IDC_EDIT1);
	DDX_Text(pDX, IDC_EDIT2, m_IDC_EDIT2);
	DDX_Text(pDX, IDC_EDIT3, m_IDC_EDIT3);
	DDX_Text(pDX, IDC_EDIT4, m_IDC_EDIT4);
	DDX_Text(pDX, IDC_EDIT5, m_IDC_EDIT5);
	DDX_Text(pDX, IDC_EDIT6, m_IDC_EDIT6);
	DDX_Text(pDX, IDC_EDIT7, m_IDC_EDIT7);
	DDX_Text(pDX, IDC_EDIT10, m_IDC_EDIT10);
	DDX_Text(pDX, IDC_EDIT15, m_IDC_EDIT15);
	DDX_Text(pDX, IDC_EDIT16, m_IDC_EDIT16);
	DDX_Text(pDX, IDC_EDIT17, m_IDC_EDIT17);
}


BEGIN_MESSAGE_MAP(XHTL_ALL, CDialogEx)
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_BUTTON2, &XHTL_ALL::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON10, &XHTL_ALL::OnBnClickedButton10)
	ON_BN_CLICKED(IDC_BUTTON1, &XHTL_ALL::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON11, &XHTL_ALL::OnBnClickedButton11)
	ON_BN_CLICKED(IDCANCEL, &XHTL_ALL::OnBnClickedCancel)
END_MESSAGE_MAP()


// XHTL_ALL 消息处理程序


BOOL XHTL_ALL::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化

	m_IDC_EDIT1=_T("Ae201901002");
	m_IDC_EDIT2=_T("路人甲");
	m_IDC_EDIT3=_T("25℃");
	m_IDC_EDIT4=_T("40%");

	 m_IDC_EDIT5=_T("20190620-20200619");
	 m_IDC_EDIT6=_T("20190620-20200619");
	 m_IDC_EDIT7=_T("20190620-20200619");
	 
	 m_IDC_EDIT10=_T("环境试验后");
	 COleDateTime t = COleDateTime::GetCurrentTime();



	 
	 m_strtime= t.Format(_T("%Y/%m/%d"));//打印填表日期




	 m_buttonclick=TRUE;




	m_Select1.AddString(_T("U-I-1"));
	m_Select1.AddString(_T("U-I-2"));
	m_Select1.AddString(_T("U-I-3"));
	m_Select1.AddString(_T("U-I-4"));
	m_Select1.SetCurSel(0);


	m_Select2.AddString(_T("U-I-1"));
	m_Select2.AddString(_T("U-I-2"));
	m_Select2.AddString(_T("U-I-3"));
	m_Select2.AddString(_T("U-I-4"));
	m_Select2.AddString(_T("I-2U-1"));
	m_Select2.AddString(_T("I-2U-2"));
	m_Select2.AddString(_T("I-2U-3"));
	m_Select2.AddString(_T("I-2U-4"));
	m_Select2.AddString(_T("I-2U-5"));
	m_Select2.AddString(_T("I-2U-6"));
	m_Select2.AddString(_T("I-2U-7"));
	m_Select2.AddString(_T("I-2U-8"));
	m_Select2.AddString(_T("I-2U-9"));
	m_Select2.AddString(_T("I-2U-10"));
	m_Select2.AddString(_T("I-2U-11"));
	m_Select2.AddString(_T("I-2U-12"));
	m_Select2.AddString(_T("I-2U-13"));
	m_Select2.AddString(_T("I-2U-14"));
	m_Select2.AddString(_T("I-2U-15"));

	m_Select2.SetCurSel(0);



	 my_Font.CreatePointFont(140, L"Times NewRoman");//创建字体和大小

	 m_Brush.CreateSolidBrush(RGB(53,203,60));
	 GetDlgItem(IDC_COMBO1)->SetFont(&my_Font);
	 GetDlgItem(IDC_COMBO2)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT14)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT8)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT9)->SetFont(&my_Font);

	 GetDlgItem(IDC_EDIT15)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT16)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT17)->SetFont(&my_Font);




	m_checkbox1=TRUE;
	m_checkbox2=FALSE;

	m_checkbox3=TRUE;
	m_checkbox4=FALSE;

	m_filenum=0;

    //GetDlgItem(IDC_62103)->EnableWindow (FALSE);//按钮不可用
	m_basename=_T("C:\\信号调理组合\\");

	strfolder=m_basename+m_externinputfilename+CString(_T("\\"));

	UpdateData(FALSE);

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


HBRUSH XHTL_ALL::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  在此更改 DC 的任何特性


	if(pWnd->GetDlgCtrlID()== IDC_COMBO1||
		 pWnd->GetDlgCtrlID()== IDC_COMBO2||
		 pWnd->GetDlgCtrlID()== IDC_EDIT8||
		 pWnd->GetDlgCtrlID()== IDC_EDIT9||
		 pWnd->GetDlgCtrlID()== IDC_EDIT14||
		 pWnd->GetDlgCtrlID()== IDC_EDIT15||
		 pWnd->GetDlgCtrlID()== IDC_EDIT16||
		 pWnd->GetDlgCtrlID()== IDC_EDIT17)
	 {
		 pDC->SetTextColor(RGB(0,0,0)); //文字颜色  
		 pDC->SetBkColor(RGB(53,203,60));
		 pDC->SetBkMode(TRANSPARENT);
		 return (HBRUSH) m_Brush.GetSafeHandle();
	 }








	// TODO:  如果默认的不是所需画笔，则返回另一个画笔
	return hbr;
}


void XHTL_ALL::OnBnClickedButton2()
{
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(TRUE);
	CString m_strpos;
	CString m_strpostemp;
	std::vector<CString> vec_strpos;
	std::vector<CString> vec_str;
	std::vector<CString> vec_Pos;
	std::vector<float> vec_float1,vec_float2,vec_float3;
	int nIndex=m_Select1.GetCurSel();
	vec_strpos.clear();
				           vec_strpos.push_back(_T("G"));
						   vec_strpos.push_back(_T("H"));
						   vec_strpos.push_back(_T("I"));
	if(m_checkbox1==TRUE||m_checkbox2==TRUE)
	{
		if(m_checkbox1==TRUE && m_checkbox2==TRUE)
		{
			MessageBox(_T("不能同时选择连接器插头621-02和621-03!!"));
			return;
		}
		
		if(m_checkbox1==TRUE)
		{

			vec_str.clear();
			vec_Pos.clear();
			vec_float1.clear();
			vec_float2.clear();
			vec_float3.clear();
			switch(nIndex)
			{

				case 0:  {
					       for(size_t j=0;j<vec_strpos.size();j++)
						   {
					           for(size_t i=5;i<10;i++)
						      {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
						   }
						   break;}
				case 1:{

					       for(size_t j=0;j<vec_strpos.size();j++)
						   {
					           for(size_t i=10;i<15;i++)
						      {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
						   }
						   break;
					   }
				case 2:{

					       for(size_t j=0;j<vec_strpos.size();j++)
						   {
					           for(size_t i=15;i<20;i++)
						      {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
						   }
						   break;
					   }
				case 3:{

					       for(size_t j=0;j<vec_strpos.size();j++)
						   {
					           for(size_t i=20;i<25;i++)
						      {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
						   }
						   break;
					   }
				default: break;
			}
		}


		if(m_checkbox2==TRUE)
		{

			vec_str.clear();
			vec_Pos.clear();
			vec_float1.clear();
			vec_float2.clear();
			vec_float3.clear();
			switch(nIndex)
			{

				case 0:  {
					       for(size_t j=0;j<vec_strpos.size();j++)
						   {
					           for(size_t i=28;i<33;i++)
						      {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
						   }
						   break;}
				case 1:{

					       for(size_t j=0;j<vec_strpos.size();j++)
						   {
					           for(size_t i=33;i<38;i++)
						      {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
						   }
						   break;
					   }
				case 2:{

					       for(size_t j=0;j<vec_strpos.size();j++)
						   {
					           for(size_t i=38;i<43;i++)
						      {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
						   }
						   break;
					   }
				case 3:{

					       for(size_t j=0;j<vec_strpos.size();j++)
						   {
					           for(size_t i=43;i<48;i++)
						      {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
						   }
						   break;
					   }
				default: break;
			}
			}

			 CString m_strtemp;
			 vec_float1=CStringtoFloat(m_IDC_EDIT8);
			 for(size_t i=0;i<vec_float1.size();i++)
			 {
				 m_strtemp.Format(_T("%.4f"),vec_float1[i]);
				 vec_str.push_back(m_strtemp);
			 }
			 vec_float2=CStringtoFloat(m_IDC_EDIT9);
			 for(size_t i=0;i<vec_float2.size();i++)
			 {
				 m_strtemp.Format(_T("%.4f"),vec_float2[i]);
				 vec_str.push_back(m_strtemp);
			 }
			for(size_t i=0;i<vec_float1.size();i++)
			{
				m_strtemp.Format(_T("%.4f"),vec_float1[i]-vec_float2[i]);
				vec_str.push_back(m_strtemp);
			}

			if(vec_str.size()!=vec_Pos.size())
			{
				MessageBox(_T("带载测量值与空载值数量不等，请君仔细检查！！"));
				return;
			}

		

			HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );
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



	}
	else
	{
		MessageBox(_T("请选择连接器插头621-02或621-03!!"));
	}





}
std::vector<float> XHTL_ALL::CStringtoFloat(CString m_s)
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

void XHTL_ALL::OnBnClickedButton10()
{
	// TODO: 在此添加控件通知处理程序代码
	if(m_buttonclick==FALSE)
	{
		MessageBox(_T("当前正在运行保存，请稍等!!"));
	}
	m_buttonclick=FALSE;
	m_filenum++;
	USES_CONVERSION;
	UpdateData(TRUE);
	CString m_strpos;
	CString m_strpostemp;
	std::vector<CString> vec_strpos;
	std::vector<CString> vec_str;
	std::vector<CString> vec_Pos;
	std::vector<float> vec_测量值,vec_x;
	CString m_txtname;
	CString m_txtoutname;

    vec_strpos.clear();
	vec_strpos.push_back(_T("G"));
	vec_strpos.push_back(_T("H"));
	vec_strpos.push_back(_T("I"));
	vec_strpos.push_back(_T("J"));

	int nIndex=m_Select2.GetCurSel();

	if(m_checkbox3==TRUE||m_checkbox4==TRUE)
	{
		if(m_checkbox3==TRUE && m_checkbox4==TRUE)
		{
			MessageBox(_T("不能同时选择连接器插头621-02和621-03!!"));
			return;
		}
		 
		vec_x.clear();
		vec_测量值.clear();

		 if(nIndex<I_2U_1)
		 {
		     for(size_t i=0;i<6;i++)
			    vec_x.push_back(i);



		 vec_测量值=CStringtoFloat(m_IDC_EDIT14);//获取数据
		 if(vec_x.size()!=vec_测量值.size())
		 {
			 MessageBox(_T("输入测量值数量不等于6，请君仔细检查!!"));
			 return;
		 }

		 if(m_checkbox3==TRUE)
		 {
			  m_txtname=strfolder+CString(_T("SourceData_62102.txt"));
			  memcpy(m_filename,(char*)T2A(m_txtname.GetBuffer(0)),100);
			  getData(m_filename,vec_x,vec_测量值);

			  m_txtoutname=strfolder+CString(_T("out_62102.txt"));
			  memcpy(m_outfilename,(char*)T2A(m_txtoutname.GetBuffer(0)),100);
			  outFileID=fopen(m_outfilename,"a+");
			  CalData(outFileID,vec_x,vec_测量值,vec_str);
			  if(vec_str.empty())
			  {
				  fclose(outFileID);
				  MessageBox(_T("拟合曲线计算模块出问题啦！，请君仔细检查!!"));
			      return;
			  }
	          fclose(outFileID);

			  switch(nIndex)
			  {
			  case 0:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=4;i<10;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J10"));
					  vec_Pos.push_back(_T("D10"));
					  vec_Pos.push_back(_T("H10"));
					  break;
				  }

			  case 1:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=11;i<17;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J17"));
					  vec_Pos.push_back(_T("D17"));
					  vec_Pos.push_back(_T("H17"));
					  break;
				  }

			  case 2:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=18;i<24;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J24"));
					  vec_Pos.push_back(_T("D24"));
					  vec_Pos.push_back(_T("H24"));
					  break;
				  }

			  case 3:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=25;i<31;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J31"));
					  vec_Pos.push_back(_T("D31"));
					  vec_Pos.push_back(_T("H31"));
					  break;
				  }




			  default: break;
			  }


		 }



		 else if(m_checkbox4==TRUE)
		 {
			  m_txtname=strfolder+CString(_T("SourceData_62103.txt"));
			  memcpy(m_filename,(char*)T2A(m_txtname.GetBuffer(0)),100);
			  getData(m_filename,vec_x,vec_测量值);

			  m_txtoutname=strfolder+CString(_T("out_62103.txt"));
			  memcpy(m_outfilename,(char*)T2A(m_txtoutname.GetBuffer(0)),100);
			  outFileID=fopen(m_outfilename,"a+");
			  CalData(outFileID,vec_x,vec_测量值,vec_str);
			  if(vec_str.empty())
			  {
				  fclose(outFileID);
				  MessageBox(_T("拟合曲线计算模块出问题啦！，请君仔细检查!!"));
			      return;
			  }
	          fclose(outFileID);

			  switch(nIndex)
			  {
			  case 0:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=4;i<10;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J10"));
					  vec_Pos.push_back(_T("D10"));
					  vec_Pos.push_back(_T("H10"));
					  break;
				  }

			  case 1:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=11;i<17;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J17"));
					  vec_Pos.push_back(_T("D17"));
					  vec_Pos.push_back(_T("H17"));
					  break;
				  }

			  case 2:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=18;i<24;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J24"));
					  vec_Pos.push_back(_T("D24"));
					  vec_Pos.push_back(_T("H24"));
					  break;
				  }

			  case 3:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=25;i<31;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J31"));
					  vec_Pos.push_back(_T("D31"));
					  vec_Pos.push_back(_T("H31"));
					  break;
				  }


			  default: break;
			  }


		 }

		 }
		 else if(nIndex>=I_2U_1)//选择I_2U测试
		 {
		     for(size_t i=1;i<6;i++)
			    vec_x.push_back(i*4);


		 vec_测量值=CStringtoFloat(m_IDC_EDIT14);//获取数据
		 if(vec_x.size()!=vec_测量值.size())
		 {
			 MessageBox(_T("输入测量值数量不等于5，请君仔细检查!!"));
			 return;
		 }

		 if(m_checkbox3==TRUE)
		 {

			     vec_strpos.clear();
				vec_strpos.push_back(_T("G"));
				vec_strpos.push_back(_T("H"));
				vec_strpos.push_back(_T("I"));
				vec_strpos.push_back(_T("J"));

			  m_txtname=strfolder+CString(_T("SourceData_I_2U_62102.txt"));
			  memcpy(m_filename,(char*)T2A(m_txtname.GetBuffer(0)),100);
			  getData(m_filename,vec_x,vec_测量值);

			  m_txtoutname=strfolder+CString(_T("out_I_2U_62102.txt"));
			  memcpy(m_outfilename,(char*)T2A(m_txtoutname.GetBuffer(0)),100);
			  outFileID=fopen(m_outfilename,"a+");
			  CalData(outFileID,vec_x,vec_测量值,vec_str);
			  if(vec_str.empty())
			  {
				  fclose(outFileID);
				  MessageBox(_T("拟合曲线计算模块出问题啦！，请君仔细检查!!"));
			      return;
			  }
	          fclose(outFileID);

			  switch(nIndex)
			  {
			  case I_2U_5:
              case I_2U_9:
			  case I_2U_13:
			  case I_2U_1:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=4;i<9;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J9"));
					  vec_Pos.push_back(_T("F9"));
					  vec_Pos.push_back(_T("I9"));
					  break;
				  }
			  case I_2U_6:
			  case I_2U_10:
			  case I_2U_14:
			  case I_2U_2:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=10;i<15;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J15"));
					  vec_Pos.push_back(_T("F15"));
					  vec_Pos.push_back(_T("I15"));
					  break;
				  }
                  case I_2U_7:
				  case I_2U_11:
				  case I_2U_15:
				  case I_2U_3:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=16;i<21;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J21"));
					  vec_Pos.push_back(_T("F21"));
					  vec_Pos.push_back(_T("I21"));
					  break;
				  }
				  case I_2U_8:
				  case I_2U_12:
				  case I_2U_4:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=22;i<27;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("J27"));
					  vec_Pos.push_back(_T("F27"));
					  vec_Pos.push_back(_T("I27"));
					  break;
				  }


				  default: break;

			  }
		 }


		 else  if(m_checkbox4==TRUE)
		 {
			     vec_strpos.clear();
				vec_strpos.push_back(_T("M"));
				vec_strpos.push_back(_T("N"));
				vec_strpos.push_back(_T("O"));
				vec_strpos.push_back(_T("P"));


			  m_txtname=strfolder+CString(_T("SourceData_I_2U_62103.txt"));
			  memcpy(m_filename,(char*)T2A(m_txtname.GetBuffer(0)),100);
			  getData(m_filename,vec_x,vec_测量值);

			  m_txtoutname=strfolder+CString(_T("out_I_2U_62103.txt"));
			  memcpy(m_outfilename,(char*)T2A(m_txtoutname.GetBuffer(0)),100);
			  outFileID=fopen(m_outfilename,"a+");
			  CalData(outFileID,vec_x,vec_测量值,vec_str);
			  if(vec_str.empty())
			  {
				  fclose(outFileID);
				  MessageBox(_T("拟合曲线计算模块出问题啦！，请君仔细检查!!"));
			      return;
			  }
	          fclose(outFileID);

			  switch(nIndex)
			  {
			  case I_2U_5:
              case I_2U_9:
			  case I_2U_13:
			  case I_2U_1:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=4;i<9;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("P9"));
					  vec_Pos.push_back(_T("L9"));
					  vec_Pos.push_back(_T("O9"));
					  
					  break;
				  }
			  case I_2U_6:
			  case I_2U_10:
			  case I_2U_14:
			  case I_2U_2:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=10;i<15;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("P15"));
					  vec_Pos.push_back(_T("L15"));
					  vec_Pos.push_back(_T("O15"));
					  
					  break;
				  }
                  case I_2U_7:
				  case I_2U_11:
				  case I_2U_15:
				  case I_2U_3:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=16;i<21;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("P21"));
					  vec_Pos.push_back(_T("L21"));
					  vec_Pos.push_back(_T("O21"));
					  
					  break;
				  }
				  case I_2U_8:
				  case I_2U_12:
				  case I_2U_4:
				  {
					  for(size_t j=0;j<vec_strpos.size();j++)
					  {
					       for(size_t i=22;i<27;i++)
						   {
					             m_strpostemp.Format(_T("%d"),i);
								 m_strpos=vec_strpos[j]+m_strpostemp;
								 vec_Pos.push_back(m_strpos);
							  }
					  }
					  vec_Pos.push_back(_T("P27"));
					  vec_Pos.push_back(_T("L27"));
					  vec_Pos.push_back(_T("O27"));
					  
					  break;
				  }

				  default: break;

			  }
		 }//endif



		 }







            HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );
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
			if(m_checkbox3==TRUE)
			{
				if(nIndex<I_2U_1)
	               sheet = sheets.get_Item(COleVariant((short)5));
				else if(nIndex>=I_2U_1 && nIndex<=I_2U_4)
				{
					sheet = sheets.get_Item(COleVariant((short)7));
				}
				else if(nIndex>=I_2U_5 && nIndex<=I_2U_8)
				{
					sheet = sheets.get_Item(COleVariant((short)8));
				}
				else if(nIndex>=I_2U_9 && nIndex<=I_2U_12)
				{
					sheet = sheets.get_Item(COleVariant((short)9));
				}
				else if(nIndex>=I_2U_3 && nIndex<=I_2U_15)
				{
					sheet = sheets.get_Item(COleVariant((short)10));
				}
			}
			else if(m_checkbox4==TRUE)
			{
				if(nIndex<I_2U_1)
	               sheet = sheets.get_Item(COleVariant((short)6));
				else if(nIndex>=I_2U_1 && nIndex<=I_2U_4)
				{
					sheet = sheets.get_Item(COleVariant((short)7));
				}
				else if(nIndex>=I_2U_5 && nIndex<=I_2U_8)
				{
					sheet = sheets.get_Item(COleVariant((short)8));
				}
				else if(nIndex>=I_2U_9 && nIndex<=I_2U_12)
				{
					sheet = sheets.get_Item(COleVariant((short)9));
				}
				else if(nIndex>=I_2U_3 && nIndex<=I_2U_15)
				{
					sheet = sheets.get_Item(COleVariant((short)10));
				}
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


	        books.Close(); 
            app.Quit();  			// 退出
	        //释放对象  
	        range.ReleaseDispatch();
	        sheet.ReleaseDispatch();
	        sheets.ReleaseDispatch();
	        book.ReleaseDispatch();
	        books.ReleaseDispatch();
	        app.ReleaseDispatch();






	}
	else
	{
		MessageBox(_T("请选择连接器插头621-02或621-03!!"));
	}


	m_buttonclick=TRUE;

}
void XHTL_ALL::getData(char * filename,std::vector<float> v1,std::vector<float> v2)
{
    //char* path = "C:\\1.txt";
	FILE* fd;
	fd=fopen(filename,"at");
	if(v1.size()==v2.size())
	{
	if(fd)
	{
		for(int i=0;i<v1.size();i++)
		{
			fprintf(fd,"%.2f,%f",v1[i],v2[i]);
			if(i!=v1.size()-1)
			{
				fprintf(fd,"\n");
			}

		}
		fprintf(fd,"\n");
		fclose(fd);
	}

	
	}

}
#define rank 1
void XHTL_ALL::CalData(FILE * fileID,std::vector<float> v1,std::vector<float> v2,std::vector<CString> &v_out)
{
	v_out.clear();//清理
	CString strtemp;
	CString strnum;


	for(size_t i=0;i<v2.size();i++)
	{
		strtemp.Format(_T("%.4f"),v2[i]);
	    v_out.push_back(strtemp);
	}//存测量值


	//double* x=(double*)malloc(sizeof(double)*v1.size());
	//double* y=(double*)malloc(sizeof(double)*v2.size());

	double atemp[2*(rank+1)]={0},b[rank+1]={0},a[rank+1][rank+1];

	int i,j,k;

	for(i=0;i<v1.size();i++)
	{
		atemp[1]+=v1[i];
		atemp[2]+=pow(v1[i],2);
#if 0
		atemp[3]+=pow(v1[i],3);
		atemp[4]+=pow(v1[i],4);
		atemp[5]+=pow(v1[i],5);
		atemp[6]+=pow(v1[i],6);
#endif
		b[0]+=v2[i];
		b[1]+=v2[i]*v1[i];
#if 0
		b[2]+=pow(v1[i],2)*v2[i];
		b[3]+=pow(v1[i],3)*v2[i];
#endif

	}
	atemp[0]=v1.size();

	for(i=0;i<rank+1;i++)
	{
		k=i;
		for(j=0;j<rank+1;j++)
			a[i][j]=atemp[k++];
	}

	//以下高斯列主消元法
	for(k=0;k<rank+1-1;k++)
	{
		int colnum=k;
		double mainelement=a[k][k];

		for(i=k;i<rank+1;i++)
		{
			if(fabs(a[i][k])>mainelement)
			{
				mainelement=fabs(a[i][k]);
				colnum=i;
			}
		}
		for(j=k;j<rank+1;j++)
		{
			double atemp1=a[k][j];
			a[k][j]=a[colnum][j];
			a[colnum][j]=atemp1;
		}

		double btemp=b[k];
		b[k]=b[colnum];
		b[colnum]=btemp;

		for(i=k+1;i<rank+1;i++)
		{
			double Mik=a[i][k]/a[k][k];
			for(j=k;j<rank+1;j++)
				a[i][j]-=Mik*a[k][j];
			b[i]-=Mik*b[k];
		}

		b[rank+1-1]/=a[rank+1-1][rank+1-1];
		for(i=rank+1-2;i>=0;i--)
		{
			double sum=0;
			for(j=i+1;j<rank+1;j++)
				sum+=a[i][j]*b[j];
			b[i]=(b[i]-sum)/a[i][i];
		}
	}

	//fprintf(fileID,"P(x)=%f%+fx%+fx^2%+fx^3\n\n",b[0],b[1],b[2],b[3]);
	fprintf(fileID,"%s[%d]!-------------\r\n\n",">>>>-----------Result of Data",m_filenum);
	strtemp.Format(_T("%s[%d]!---------------\r\n\n"),_T("---------------Result of Data"),m_filenum);
	//EditShowData(strtemp);

	for(i=0;i<v1.size();i++)
	{
		fprintf(fileID,"%s%d=%f,detaY=%f\r\n","Fitted Value NO.",i,v1[i]*b[1]+b[0],v2[i]-(v1[i]*b[1]+b[0]));
		strtemp.Format(_T("%s%d=%f,detaY=%f\r\n"),_T("Fitted Value NO."),i,v1[i]*b[1]+b[0],v2[i]-(v1[i]*b[1]+b[0]));
		
		strtemp.Format(_T("%.4f"),v1[i]*b[1]+b[0]);
	    v_out.push_back(strtemp);     	
	
	}
	for(i=0;i<v1.size();i++)
	{
		strtemp.Format(_T("%.4f"),v2[i]-(v1[i]*b[1]+b[0]));
	    v_out.push_back(strtemp);     	
	}
	for(size_t i=0;i<v1.size()+1;i++)
	{
		v_out.push_back(_T("√"));
	}



	fprintf(fileID,"%s\r\n","-----------------------------------");
	fprintf(fileID,"Fitted curve function:y=%f%+fx%\n\n",b[0],b[1]);
	strtemp.Format(_T("%s\r\n"),_T("-----------------------------------"));
   // EditShowData(strtemp);
	CString outstr;
	if(b[0]>=0 && b[1]>=0)
	{
		strtemp.Format(_T("Fitted curve function:y=%f+%fx\n\n"),b[0],b[1]);
		outstr.Format(_T("y=%.4f+%.4fx"),b[0],b[1]);
	}
	else if(b[0]>=0 && b[1]<0)
	{
		strtemp.Format(_T("Fitted curve function:y=%f%fx\n\n"),b[0],b[1]);
		outstr.Format(_T("y=%.4f%.4fx"),b[0],b[1]);
	}
	else if(b[0]<0 && b[1]<0)
	{
		strtemp.Format(_T("Fitted curve function:y=%f%fx\n\n"),b[0],b[1]);
		outstr.Format(_T("y=%.4f%.4fx"),b[0],b[1]);
	}
	else if(b[0]<0 && b[1]>=0)
	{
		strtemp.Format(_T("Fitted curve function:y=%f+%fx\n\n"),b[0],b[1]);
		outstr.Format(_T("y=%.4f+%.4fx"),b[0],b[1]);
	}
  //  EditShowData(strtemp);

	 v_out.push_back(outstr);//将拟合公式压入缓存

	std::vector<double> detay;

	for(i=0;i<v1.size();i++)
	{
		detay.push_back(abs(b[1]*v1[i]+b[0]-v2[i])/sqrt(b[1]*b[1]+b[0]*b[0]));
	}

	fprintf(fileID,"Linearity=%f\r\n",(*max_element(detay.begin(),detay.end()))/(*max_element(v2.begin(),v2.end())));//计算线性度
	fprintf(fileID,"%s\r\n","-----------------------------------<<<<");
	strtemp.Format(_T("    Linearity=%f\r\n"),(*max_element(detay.begin(),detay.end()))/(*max_element(v2.begin(),v2.end())));//计算线性度
    //EditShowData(strtemp);
	strtemp.Format(_T("%s\r\n"),_T("----------------------------------------------------"));
   // EditShowData(strtemp);
	strtemp.Format(_T("%.4f"),(*max_element(detay.begin(),detay.end()))/(*max_element(v2.begin(),v2.end())));
	v_out.push_back(strtemp);//计算线性度 
	

 
}


void XHTL_ALL::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(TRUE);
	std::vector<CString> vec_str;
	std::vector<CString> vec_Pos;
	 HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );
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

	         sheet = sheets.get_Item(COleVariant((short)1));

			 vec_str.push_back(m_IDC_EDIT1);
			 vec_str.push_back(m_IDC_EDIT2);
			 vec_str.push_back(m_strtime);
			 vec_str.push_back(_T("北京航天测控技术有限公司"));
			 vec_str.push_back(_T("合格（  √ ） /  不合格（   ）"));
			 vec_str.push_back(m_IDC_EDIT3);
			 vec_str.push_back(m_IDC_EDIT4);
			 vec_str.push_back(m_IDC_EDIT5);
			 vec_str.push_back(m_IDC_EDIT6);
			 vec_str.push_back(m_IDC_EDIT7);

			 vec_Pos.push_back(_T("D5"));
			 vec_Pos.push_back(_T("D6"));
			 vec_Pos.push_back(_T("D7"));
			 vec_Pos.push_back(_T("D8"));
			 vec_Pos.push_back(_T("D9"));
			 vec_Pos.push_back(_T("D10"));
			 vec_Pos.push_back(_T("D11"));
			 vec_Pos.push_back(_T("K13"));
			 vec_Pos.push_back(_T("K14"));
			 vec_Pos.push_back(_T("K15"));
			 

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






}


void XHTL_ALL::OnBnClickedButton11()
{
	// TODO: 在此添加控件通知处理程序代码


	UpdateData(TRUE);
	std::vector<CString> vec_str;
	std::vector<CString> vec_Pos;
	std::vector<float> vec_float;
	 std::vector<CString> _str_temp;
	 HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );
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

	         CString str_temp; 
	         vec_float=CStringtoFloat(m_IDC_EDIT15);//
			 if(vec_float.empty())
			 {
				  AfxMessageBox(L"输入不能为空，请君仔细检查!");
				  goto GameOver;

			 }
	         for(size_t i=0;i<vec_float.size();i++)
			 {
				 str_temp.Format(_T("%.4f"),vec_float[i]);
				 vec_str.push_back(str_temp);
			 }
			 for(size_t i=0;i<vec_float.size();i++)
			 {
				 vec_str.push_back(_T("√"));
			 }

			 vec_float=CStringtoFloat(m_IDC_EDIT16);//
			 if(vec_float.empty())
			 {
				  AfxMessageBox(L"输入不能为空，请君仔细检查!");
				  goto GameOver;

			 }
	         for(size_t i=0;i<vec_float.size();i++)
			 {
				 str_temp.Format(_T("%.4f"),vec_float[i]);
				 vec_str.push_back(str_temp);
			 }

			 for(size_t i=0;i<vec_float.size();i++)
			 {
				 vec_str.push_back(_T("√"));
			 }

			 vec_float=CStringtoFloat(m_IDC_EDIT17);//

			 if(vec_float.empty())
			 {
				  AfxMessageBox(L"输入不能为空，请君仔细检查!");
				  goto GameOver;

			 }

	         for(size_t i=0;i<vec_float.size();i++)
			 {
				 str_temp.Format(_T("%.4f"),vec_float[i]);
				 vec_str.push_back(str_temp);
			 }

			 for(size_t i=0;i<vec_float.size()-1;i++)
			 {
				 vec_str.push_back(_T("√"));
			 }


			
			 _str_temp.push_back(_T("E"));
			 _str_temp.push_back(_T("F"));
			 


			 for(size_t i=0;i<_str_temp.size();i++)
			 {
				for(size_t j=4;j<12;j++)
				{
					str_temp.Format(_T("%d"),j);
					str_temp=_str_temp[i]+str_temp;
					vec_Pos.push_back(str_temp);
				}
			 }

			 for(size_t i=0;i<_str_temp.size();i++)
			 {
				for(size_t j=16;j<31;j++)
				{
					str_temp.Format(_T("%d"),j);
					str_temp=_str_temp[i]+str_temp;
					vec_Pos.push_back(str_temp);
				}
			 }

			 
			 vec_Pos.push_back(_T("C35"));
			 
			 vec_Pos.push_back(_T("E35"));
			 vec_Pos.push_back(_T("F35"));
			if(vec_str.size()!=vec_Pos.size())
			{
				  AfxMessageBox(L"输入数据数量不正确，请君仔细检查!");
				  goto GameOver;
			}


	        //得到Workbook
	        book.AttachDispatch(lpDisp);
	        //得到Worksheets
	        sheets.AttachDispatch(book.get_Worksheets());

	         sheet = sheets.get_Item(COleVariant((short)3));

			



			 

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



}


void XHTL_ALL::OnBnClickedCancel()
{
	// TODO: 在此添加控件通知处理程序代码

	if(AfxMessageBox(_T("您确定要退出当前信号调理单机填表工作，劝君三思而行!"),MB_YESNO|MB_ICONEXCLAMATION)==IDYES)
	{
		CDialogEx::OnCancel();
	}

}

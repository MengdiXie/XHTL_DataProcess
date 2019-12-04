// XHTL.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "input.h"
#include "XHTL.h"
#include "afxdialogex.h"


// XHTL �Ի���
extern CString exfilename;//�ⲿ����
extern CString m_inputopname;
#define WM_UPDATE_MESSAGE2 10006
#define WM_UPDATE_MESSAGE 10005
IMPLEMENT_DYNAMIC(XHTL, CDialogEx)

XHTL::XHTL(CWnd* pParent /*=NULL*/)
	: CDialogEx(XHTL::IDD, pParent)
{

}

XHTL::~XHTL()
{
}

void XHTL::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);


	DDX_Text(pDX, IDC_EDIT1, m_��Ʒ����);
	DDX_Text(pDX, IDC_EDIT2, m_boardnum);

	DDX_Text(pDX, IDC_EDIT3, m_opname);
	DDX_Text(pDX, IDC_EDIT4, m_tempture);

	DDX_Text(pDX, IDC_EDIT5, m_humidity);
	DDX_Text(pDX, IDC_EDIT6, m_multimeter);

	DDX_Text(pDX, IDC_EDIT7, m_constant_volt_source);
	DDX_Text(pDX, IDC_EDIT8, m_signal_source);

	DDX_Text(pDX, IDC_EDIT9, m_���Թ�װ);

	DDX_Text(pDX, IDC_EDIT10, m_EDIT10);
	DDX_Text(pDX, IDC_EDIT14, m_EDIT14);
	DDX_Text(pDX, IDC_EDIT15, m_EDIT15);
	DDX_Text(pDX, IDC_EDIT16, m_EDIT16);

	DDX_Text(pDX, IDC_EDIT17, m_EDIT17);
	DDX_Text(pDX, IDC_EDIT18, m_EDIT18);
	DDX_Text(pDX, IDC_EDIT19, m_EDIT19);
	DDX_Text(pDX, IDC_EDIT20, m_EDIT20);
	DDX_Text(pDX, IDC_EDIT21, m_EDIT21);
	DDX_Text(pDX, IDC_EDIT22, m_EDIT22);
}


BEGIN_MESSAGE_MAP(XHTL, CDialogEx)
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_BUTTON1, &XHTL::OnBnClickedButton1)
	ON_MESSAGE(WM_UPDATE_MESSAGE, OnUpdateTrueMessage)
	ON_BN_CLICKED(IDC_BUTTON3, &XHTL::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &XHTL::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON5, &XHTL::OnBnClickedButton5)
END_MESSAGE_MAP()


// XHTL ��Ϣ�������


HBRUSH XHTL::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  �ڴ˸��� DC ���κ�����

		 if(pWnd->GetDlgCtrlID()== IDC_EDIT10||
		 pWnd->GetDlgCtrlID()== IDC_EDIT14||
		 pWnd->GetDlgCtrlID()== IDC_EDIT15||
		 pWnd->GetDlgCtrlID()== IDC_EDIT16||
		 pWnd->GetDlgCtrlID()== IDC_EDIT17||
		 pWnd->GetDlgCtrlID()== IDC_EDIT18||
		 pWnd->GetDlgCtrlID()== IDC_EDIT19||
		 pWnd->GetDlgCtrlID()== IDC_EDIT20||
		 pWnd->GetDlgCtrlID()== IDC_EDIT21||
		 pWnd->GetDlgCtrlID()== IDC_EDIT22)
	 {
		 pDC->SetTextColor(RGB(0,0,0)); //������ɫ  
		 pDC->SetBkColor(RGB(53,203,60));
		 pDC->SetBkMode(TRANSPARENT);
		 return (HBRUSH) m_Brush.GetSafeHandle();
	 }


	// TODO:  ���Ĭ�ϵĲ������軭�ʣ��򷵻���һ������
	return hbr;
}


BOOL XHTL::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��

	m_��Ʒ����=_T("1XK621-5X");
	m_boardnum=_T("�źŵ������Ae201901002");
	m_opname=_T("·�˼�");
	m_tempture=_T("25��");
	m_humidity=_T("40%");

	m_multimeter=_T("20190620-20200619");
	m_constant_volt_source=_T("20190620-20200619");
	m_signal_source=_T("20190620-20200619");
	m_���Թ�װ=_T("20190620-20200619");
	 
    m_click=TRUE;//��ʼ���Ե���


	COleDateTime t = COleDateTime::GetCurrentTime();

	m_strtime=t.Format(_T("%Y/%m/%d"));//��ӡ�������


	 my_Font.CreatePointFont(140, L"Times NewRoman");//��������ʹ�С

	 m_Brush.CreateSolidBrush(RGB(53,203,60));


	 GetDlgItem(IDC_EDIT10)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT14)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT15)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT16)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT17)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT18)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT19)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT20)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT21)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT22)->SetFont(&my_Font);

	UpdateData(FALSE);

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}

LRESULT XHTL::OnUpdateTrueMessage(WPARAM wParam, LPARAM lParam)
{
	BOOL f=(BOOL)wParam;
	if(f==TRUE)
       UpdateData(TRUE);
	else
	   UpdateData(FALSE);
    return 0;
}



void XHTL::OnBnClickedButton1()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_click==TRUE)
	{
		m_click=FALSE;

	    AfxBeginThread((AFX_THREADPROC)XHTL::Button1Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}

}

DWORD XHTL::Button1Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	XHTL *pMain=(XHTL*) lpParam;	

	std::vector<CString> vec_str;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!(pMain->app.CreateDispatch(L"Excel.Application")))
	{
	    AfxMessageBox(L"�޷�����Excel������!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);

	
	
	//�õ�Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//�õ�Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)1));




	pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);
	vec_str.push_back(pMain->m_��Ʒ����);
	vec_str.push_back(pMain->m_boardnum);
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	//vec_str.push_back(_T("D"));
	vec_str.push_back(_T("�ϸ�  ��  �� /  ���ϸ�   ��"));
	vec_str.push_back(pMain->m_tempture);
	vec_str.push_back(pMain->m_humidity);
	vec_str.push_back(pMain->m_multimeter);
	vec_str.push_back(pMain->m_constant_volt_source);
	vec_str.push_back(pMain->m_signal_source);
	vec_str.push_back(pMain->m_���Թ�װ);

	

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
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	//pMain->app.put_Visible(TRUE);
	//pMain->book.Save();
	//pMain->book.put_Saved(TRUE);



	

	pMain->books.Close(); 
	
    pMain->app.Quit();  			// �˳�
	//�ͷŶ���  
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


void XHTL::OnBnClickedButton3()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_click==TRUE)
	{
		m_click=FALSE;

	    AfxBeginThread((AFX_THREADPROC)XHTL::Button2Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}

}

std::vector<float> XHTL::CStringtoFloat(CString m_s)
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
DWORD XHTL::Button2Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	XHTL *pMain=(XHTL*) lpParam;	

   pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);
	m_inputopname=pMain->m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!pMain->app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"�޷�����Excel������!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//�õ�Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//�õ�Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)2));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_EDIT10);//

	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto XHTL_GameOver1;

		return TRUE;
	}
	
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	vec_float=pMain->CStringtoFloat(pMain->m_EDIT14);//

	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto XHTL_GameOver1;

		return TRUE;
	}
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	vec_Pos.push_back(_T("G4"));
	vec_Pos.push_back(_T("G5"));
	vec_Pos.push_back(_T("G6"));
	vec_Pos.push_back(_T("G7"));
	vec_Pos.push_back(_T("G8"));
	vec_Pos.push_back(_T("G9"));
	vec_Pos.push_back(_T("G10"));
	vec_Pos.push_back(_T("G11"));

	vec_Pos.push_back(_T("G17"));
	vec_Pos.push_back(_T("G18"));
	vec_Pos.push_back(_T("G19"));
	vec_Pos.push_back(_T("G20"));
	vec_Pos.push_back(_T("G21"));
	vec_Pos.push_back(_T("G22"));
	vec_Pos.push_back(_T("G23"));
	vec_Pos.push_back(_T("G24"));

	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("������С��16������!!"));

		goto XHTL_GameOver1;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp =pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);


XHTL_GameOver1:
	    pMain->books.Close(); 
        pMain->app.Quit();  			// �˳�
	    //�ͷŶ���  
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
#if 1
void XHTL::OnBnClickedButton4()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
#if 1
	if(m_click==TRUE)
	{
		m_click=FALSE;

	    AfxBeginThread((AFX_THREADPROC)XHTL::Button3Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}
#endif
}
DWORD XHTL::Button3Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	XHTL *pMain=(XHTL*) lpParam;	

   pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);
	m_inputopname=pMain->m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!pMain->app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"�޷�����Excel������!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//�õ�Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//�õ�Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)3));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_EDIT15);//
	if(vec_float.empty())
	{
		AfxMessageBox(_T("EDIT15���벻��Ϊ�գ�����!!"));
		goto XHTL_GameOver2;
		return TRUE;
	}
	
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	vec_float=pMain->CStringtoFloat(pMain->m_EDIT16);//

	if(vec_float.empty())
	{
		AfxMessageBox(_T("EDIT16���벻��Ϊ�գ�����!!"));

		goto XHTL_GameOver2;

		return TRUE;
	}
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ


	vec_float=pMain->CStringtoFloat(pMain->m_EDIT17);//
	if(vec_float.empty())
	{
		AfxMessageBox(_T("EDIT17���벻��Ϊ�գ�����!!"));
		goto XHTL_GameOver2;
		return TRUE;
	}
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	vec_float=pMain->CStringtoFloat(pMain->m_EDIT18);//
	if(vec_float.empty())
	{
		AfxMessageBox(_T("EDIT18���벻��Ϊ�գ�����!!"));
		goto XHTL_GameOver2;
		return TRUE;
	}
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ



	vec_Pos.push_back(_T("G4"));
	vec_Pos.push_back(_T("G5"));
	vec_Pos.push_back(_T("G6"));
	vec_Pos.push_back(_T("G7"));
	vec_Pos.push_back(_T("G8"));
	vec_Pos.push_back(_T("G9"));
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


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("������С��30������!!"));

		goto XHTL_GameOver2;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp =pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);


XHTL_GameOver2:
	    pMain->books.Close(); 
        pMain->app.Quit();  			// �˳�
	    //�ͷŶ���  
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



void XHTL::OnBnClickedButton5()
{
#if 1
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_click==TRUE)
	{
		m_click=FALSE;

	    AfxBeginThread((AFX_THREADPROC)XHTL::Button4Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}
#endif
}
#endif


DWORD XHTL::Button4Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	XHTL *pMain=(XHTL*) lpParam;	

   pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);
	m_inputopname=pMain->m_opname;
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!pMain->app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"�޷�����Excel������!");
	    return TRUE;
	}
	pMain->books.AttachDispatch(pMain->app.get_Workbooks());
	pMain->lpDisp = pMain->books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);
	
	//�õ�Workbook
	pMain->book.AttachDispatch(pMain->lpDisp);
	//�õ�Worksheets
	pMain->sheets.AttachDispatch(pMain->book.get_Worksheets());

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)3));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_EDIT19);//
	if(vec_float.empty())
	{
		AfxMessageBox(_T("EDIT19���벻��Ϊ�գ�����!!"));
		goto XHTL_GameOver3;
		return TRUE;
	}
	
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	vec_float=pMain->CStringtoFloat(pMain->m_EDIT20);//

	if(vec_float.empty())
	{
		AfxMessageBox(_T("EDIT20���벻��Ϊ�գ�����!!"));

		goto XHTL_GameOver3;

		return TRUE;
	}
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ


	vec_float=pMain->CStringtoFloat(pMain->m_EDIT21);//
	if(vec_float.empty())
	{
		AfxMessageBox(_T("EDIT21���벻��Ϊ�գ�����!!"));
		goto XHTL_GameOver3;
		return TRUE;
	}
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	vec_float=pMain->CStringtoFloat(pMain->m_EDIT22);//
	if(vec_float.empty())
	{
		AfxMessageBox(_T("EDIT22���벻��Ϊ�գ�����!!"));
		goto XHTL_GameOver3;
		return TRUE;
	}
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ



	vec_Pos.push_back(_T("J4"));
	vec_Pos.push_back(_T("J5"));
	vec_Pos.push_back(_T("J6"));
	vec_Pos.push_back(_T("J7"));
	vec_Pos.push_back(_T("J8"));
	vec_Pos.push_back(_T("J9"));
	vec_Pos.push_back(_T("J10"));
	vec_Pos.push_back(_T("J11"));


	vec_Pos.push_back(_T("J12"));
	vec_Pos.push_back(_T("J13"));
	vec_Pos.push_back(_T("J14"));
	vec_Pos.push_back(_T("J15"));
	vec_Pos.push_back(_T("J16"));
	vec_Pos.push_back(_T("J17"));
	vec_Pos.push_back(_T("J18"));
	vec_Pos.push_back(_T("J19"));

	vec_Pos.push_back(_T("J20"));
	vec_Pos.push_back(_T("J21"));
	vec_Pos.push_back(_T("J22"));
	vec_Pos.push_back(_T("J23"));
	vec_Pos.push_back(_T("J24"));
	vec_Pos.push_back(_T("J25"));
	vec_Pos.push_back(_T("J26"));
	vec_Pos.push_back(_T("J27"));


	vec_Pos.push_back(_T("J28"));
	vec_Pos.push_back(_T("J29"));
	vec_Pos.push_back(_T("J30"));
	vec_Pos.push_back(_T("J31"));
	vec_Pos.push_back(_T("J32"));
	vec_Pos.push_back(_T("J33"));


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("������С��30������!!"));

		goto XHTL_GameOver3;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp =pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);


XHTL_GameOver3:
	    pMain->books.Close(); 
        pMain->app.Quit();  			// �˳�
	    //�ͷŶ���  
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
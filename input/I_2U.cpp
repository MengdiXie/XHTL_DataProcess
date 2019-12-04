// I_2U.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "input.h"
#include "I_2U.h"
#include "afxdialogex.h"
#include <afxdisp.h> 

extern CString exfilename;//�ⲿ����
extern CString m_inputopname;
extern BOOL m_closeornot;//�ж϶Ի����Ƿ�ر�
// I_2U �Ի���

#define WM_UPDATE_MESSAGE2 10002
#define WM_UPDATE_MESSAGE 10001
IMPLEMENT_DYNAMIC(I_2U, CDialogEx)

I_2U::I_2U(CWnd* pParent /*=NULL*/)
	: CDialogEx(I_2U::IDD, pParent)
{

}

I_2U::~I_2U()
{
}

void I_2U::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Text(pDX, IDC_EDIT1, m_boardnum);
	DDX_Text(pDX, IDC_EDIT2, m_opname);

	DDX_Text(pDX, IDC_EDIT3, m_tempture);
	DDX_Text(pDX, IDC_EDIT4, m_humidity);

	DDX_Text(pDX, IDC_EDIT5, m_multimeter);
	DDX_Text(pDX, IDC_EDIT6, m_constant_volt_source);

	DDX_Text(pDX, IDC_EDIT7, m_signal_source);
	DDX_Text(pDX, IDC_EDIT8, m_���Թ�װ);

	DDX_Text(pDX, IDC_EDIT10, m_���Խ׶�);

	DDX_Text(pDX, IDC_tab2, m_tab2);
	DDX_Text(pDX, IDC_tab3, m_tab3);
	DDX_Text(pDX, IDC_tab4, m_tab4);
	DDX_Text(pDX, IDC_tab5, m_tab5);
	DDX_Text(pDX, IDC_tab6, m_tab6);
	DDX_Text(pDX, IDC_tab7, m_tab7);
	DDX_Text(pDX, IDC_tab8, m_tab8);

	DDX_Text(pDX, IDC_tab9, m_tab9);
	DDX_Text(pDX, IDC_tab10, m_tab10);
	DDX_Text(pDX, IDC_tab11, m_tab11);
	DDX_Text(pDX, IDC_tab12, m_tab12);

	
}


BEGIN_MESSAGE_MAP(I_2U, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &I_2U::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &I_2U::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, &I_2U::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &I_2U::OnBnClickedButton4)
	ON_MESSAGE(WM_UPDATE_MESSAGE, OnUpdateTrueMessage)
	ON_MESSAGE(WM_UPDATE_MESSAGE2, OnUpdateFalseMessage)
	ON_BN_CLICKED(IDC_BUTTON5, &I_2U::OnBnClickedButton5)
	ON_BN_CLICKED(IDC_BUTTON6, &I_2U::OnBnClickedButton6)
	ON_BN_CLICKED(IDC_BUTTON7, &I_2U::OnBnClickedButton7)
	ON_BN_CLICKED(IDC_BUTTON8, &I_2U::OnBnClickedButton8)
	ON_BN_CLICKED(IDC_BUTTON9, &I_2U::OnBnClickedButton9)
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDCANCEL, &I_2U::OnBnClickedCancel)
	ON_BN_CLICKED(IDOK, &I_2U::OnBnClickedOk)
END_MESSAGE_MAP()


// I_2U ��Ϣ�������


BOOL I_2U::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��


	m_boardnum=_T("I_2U_201901001");
	m_opname=_T("·�˼�");
	m_tempture=_T("25��");
	m_humidity=_T("40%");

	 m_multimeter=_T("20190620-20200619");
	 m_constant_volt_source=_T("20190620-20200619");
	 m_signal_source=_T("20190620-20200619");
	 m_���Թ�װ=_T("20190620-20200619");
	 m_���Խ׶�=_T("���廷�������");

	 m_tab3=_T("8,8,8,8,8,8,8,8,8,8,8,8");
	 m_tab4=_T("250,8,8,8,8,8,8,8,8,8,8,8,8,250,8,8");
	 m_tab5=_T("8,8,8,8,8,8,8,8,250,8,8,8,8,8,8,8,8");
	 m_tab6=_T("1.68,1.68,50,8,8,8,8,1.68,50,8,8,8,8,50,8,8,8");

	 COleDateTime t = COleDateTime::GetCurrentTime();
	 m_OKorNot=_T("�ϸ�� ���ϸ�");

	 my_Font.CreatePointFont(140, L"Times NewRoman");//��������ʹ�С

	 m_Brush.CreateSolidBrush(RGB(53,203,60));
	 
	 m_strtime= t.Format(_T("%Y/%m/%d"));//��ӡ�������

	 GetDlgItem(IDC_tab2)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab3)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab4)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab5)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab6)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab7)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab8)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab9)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab10)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab11)->SetFont(&my_Font);
	 GetDlgItem(IDC_tab12)->SetFont(&my_Font); 

	 m_click=TRUE;//��ʼ���Ե���


	UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}
LRESULT I_2U::OnUpdateTrueMessage(WPARAM wParam, LPARAM lParam)
{
	BOOL f=(BOOL)wParam;
	if(f==TRUE)
       UpdateData(TRUE);
	else
	   UpdateData(FALSE);
    return 0;
}
LRESULT I_2U::OnUpdateFalseMessage(WPARAM wParam, LPARAM lParam)
{
    UpdateData(FALSE);
    return 0;
}

void I_2U::OnBnClickedButton1()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_click==TRUE)
	{
		m_click=FALSE;
#if 0
	std::vector<CString> vec_str;
	std::vector<CString> vec_Pos;
	COleVariant vResult;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if(!app.CreateDispatch(L"Excel.Application"))
	{
	    AfxMessageBox(L"�޷�����Excel������!");
	    return;
	}
	books.AttachDispatch(app.get_Workbooks());
	lpDisp = books.Open(exfilename,covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional);

	
	
	//�õ�Workbook
	book.AttachDispatch(lpDisp);
	//�õ�Worksheets
	sheets.AttachDispatch(book.get_Worksheets());

	

	sheet = sheets.get_Item(COleVariant((short)1));

	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//��ӡ�������


	UpdateData(TRUE);
	vec_str.push_back(m_boardnum);
	vec_str.push_back(m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("D"));
	vec_str.push_back(_T("�ϸ�  ��  �� /  ���ϸ�   ��"));
	vec_str.push_back(m_tempture);
	vec_str.push_back(m_humidity);
	vec_str.push_back(m_multimeter);
	vec_str.push_back(m_constant_volt_source);
	vec_str.push_back(m_signal_source);
	vec_str.push_back(m_���Թ�װ);

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
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	books.Close(); 
    app.Quit();  			// �˳�
	//�ͷŶ���  
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	app.ReleaseDispatch();
#else
	AfxBeginThread((AFX_THREADPROC)I_2U::Button1Thread,this,THREAD_PRIORITY_IDLE);
#endif
	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}

}
std::vector<float> I_2U::CStringtoFloat(CString m_s)
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

void I_2U::OnBnClickedButton2()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_click==TRUE)
	{
		m_click=FALSE;
	    AfxBeginThread((AFX_THREADPROC)I_2U::Button2Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}
}


void I_2U::OnBnClickedButton3()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

		// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_click==TRUE)
	{
		m_click=FALSE;
	    AfxBeginThread((AFX_THREADPROC)I_2U::Button3Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}
}


void I_2U::OnBnClickedButton4()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_click==TRUE)
	{
		m_click=FALSE;
	    AfxBeginThread((AFX_THREADPROC)I_2U::Button4Thread,this,THREAD_PRIORITY_IDLE);

	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}
}


DWORD I_2U::Button1Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

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

	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//��ӡ�������


	pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);
	vec_str.push_back(pMain->m_boardnum);
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(strtime);
	vec_str.push_back(_T("D"));
	vec_str.push_back(_T("�ϸ�  ��  �� /  ���ϸ�   ��"));
	vec_str.push_back(pMain->m_tempture);
	vec_str.push_back(pMain->m_humidity);
	vec_str.push_back(pMain->m_multimeter);
	vec_str.push_back(pMain->m_constant_volt_source);
	vec_str.push_back(pMain->m_signal_source);
	vec_str.push_back(pMain->m_���Թ�װ);
	vec_str.push_back(pMain->m_���Խ׶�);
	

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
    vec_Pos.push_back(_T("B7"));

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

DWORD I_2U::Button2Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

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
	vec_float=pMain->CStringtoFloat(pMain->m_tab2);//

	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto GameOver3;

		return TRUE;
	}
	
	for(size_t i=0;i<vec_float.size();i++)
	{
		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	for(size_t i=0;i<16;i++)
	{
		vec_str.push_back(_T("��"));
	}
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);


	vec_Pos.push_back(_T("F12"));
	vec_Pos.push_back(_T("F13"));
	vec_Pos.push_back(_T("F14"));
	vec_Pos.push_back(_T("F15"));
	vec_Pos.push_back(_T("F16"));
	vec_Pos.push_back(_T("F17"));
	vec_Pos.push_back(_T("F18"));
	vec_Pos.push_back(_T("F19"));
	vec_Pos.push_back(_T("F20"));
	vec_Pos.push_back(_T("F21"));
	vec_Pos.push_back(_T("F22"));
	vec_Pos.push_back(_T("F23"));

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
	vec_Pos.push_back(_T("F3"));
	vec_Pos.push_back(_T("F4"));
	vec_Pos.push_back(_T("F5"));
	vec_Pos.push_back(_T("F6"));



	vec_Pos.push_back(_T("B24"));
	vec_Pos.push_back(_T("E24"));
	vec_Pos.push_back(_T("H24"));

	vec_Pos.push_back(_T("B7"));
	vec_Pos.push_back(_T("E7"));
	vec_Pos.push_back(_T("H7"));

	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("������С��12������!!"));

		goto GameOver3;

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


GameOver3:
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

DWORD I_2U::Button3Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

	pMain->PostMessage(WM_UPDATE_MESSAGE,TRUE,0);
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

	vec_float=pMain->CStringtoFloat(pMain->m_tab3);//

	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto GameOver4;

		return TRUE;
	}


	
	for(size_t i=0;i<vec_float.size();i++)
	{
		if((int)vec_float[i]==8)
		{
			str_temp=_T("��");
		}
		else
		{
			str_temp.Format(_T("%.4f"),vec_float[i]);
		}
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	for(size_t i=0;i<12;i++)
	{
		vec_str.push_back(_T("��"));
	}
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);




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

	vec_Pos.push_back(_T("B16"));
	vec_Pos.push_back(_T("E16"));
	vec_Pos.push_back(_T("H16"));


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("������С��12������!"));

		goto GameOver4;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);


GameOver4:
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

DWORD I_2U::Button4Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

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

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)4));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_tab4);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto GameOver5;

		return TRUE;
	}


	for(size_t i=0;i<vec_float.size();i++)
	{
		if((int)vec_float[i]==8)
		{
			str_temp=_T("��");
		}
		else
		{
			str_temp.Format(_T("%.4f"),vec_float[i]);
		}
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	for(size_t i=0;i<16;i++)
	{
		vec_str.push_back(_T("��"));
	}
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);


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

	vec_Pos.push_back(_T("B20"));
	vec_Pos.push_back(_T("E20"));
	vec_Pos.push_back(_T("H20"));

	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("������С��16������!"));

		goto GameOver5;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}


	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);

GameOver5:
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

void I_2U::OnBnClickedButton5()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_click==TRUE)
	{
	   m_click=FALSE;
	   AfxBeginThread((AFX_THREADPROC)I_2U::Button5Thread,this,THREAD_PRIORITY_IDLE);
	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}

}


DWORD I_2U::Button5Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

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

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)5));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_tab5);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto GameOver6;

		return TRUE;
	}


	for(size_t i=0;i<vec_float.size();i++)
	{
		if((int)vec_float[i]==8)
		{
			str_temp=_T("��");
		}
		else
		{
			str_temp.Format(_T("%.4f"),vec_float[i]);
		}
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	for(size_t i=0;i<17;i++)
	{
		vec_str.push_back(_T("��"));
	}

	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);


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

	vec_Pos.push_back(_T("B21"));
	vec_Pos.push_back(_T("E21"));
	vec_Pos.push_back(_T("H21"));

	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("������С��17������!"));

		goto GameOver6;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);

GameOver6:
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

void I_2U::OnBnClickedButton6()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	if(m_click==TRUE)
	{
	   m_click=FALSE;
	   AfxBeginThread((AFX_THREADPROC)I_2U::Button6Thread,this,THREAD_PRIORITY_IDLE);
	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}
}



DWORD I_2U::Button6Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

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

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)6));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_tab6);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto GameOver7;

		return TRUE;
	}

	for(size_t i=0;i<vec_float.size();i++)
	{
		if((int)vec_float[i]==8)
		{
			str_temp=_T("��");
		}
		else
		{
			str_temp.Format(_T("%.4f"),vec_float[i]);
		}
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	for(size_t i=0;i<17;i++)
	{
		vec_str.push_back(_T("��"));
	}

	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);


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

	vec_Pos.push_back(_T("B21"));
	vec_Pos.push_back(_T("E21"));
	vec_Pos.push_back(_T("H21"));

	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("������С��17������!"));

		goto GameOver7;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);

GameOver7:
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

void I_2U::OnBnClickedButton7()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	
	if(m_click==TRUE)
	{
	   m_click=FALSE;
	   AfxBeginThread((AFX_THREADPROC)I_2U::Button7Thread,this,THREAD_PRIORITY_IDLE);
	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}

}

DWORD I_2U::Button7Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

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

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)7));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_tab7);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto GameOver9;

		return TRUE;
	}


	for(size_t i=0;i<vec_float.size();i++)
	{

		if((int)vec_float[i]==8)
		{
			str_temp=_T("��");
		}
		else
		{
			str_temp.Format(_T("%.4f"),vec_float[i]);
		}
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	vec_float=pMain->CStringtoFloat(pMain->m_tab8);//
	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ
	
	
	for(size_t i=0;i<10;i++)
	{
		vec_str.push_back(_T("��"));
	}

	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);


	vec_Pos.push_back(_T("H4"));
	vec_Pos.push_back(_T("H5"));
	vec_Pos.push_back(_T("H6"));
	vec_Pos.push_back(_T("H7"));
	vec_Pos.push_back(_T("H8"));
	vec_Pos.push_back(_T("H9"));
	vec_Pos.push_back(_T("H10"));

	vec_Pos.push_back(_T("H17"));
	vec_Pos.push_back(_T("H18"));
	vec_Pos.push_back(_T("H19"));	

	vec_Pos.push_back(_T("I4"));
	vec_Pos.push_back(_T("I5"));
	vec_Pos.push_back(_T("I6"));
	vec_Pos.push_back(_T("I7"));
	vec_Pos.push_back(_T("I8"));
	vec_Pos.push_back(_T("I9"));
	vec_Pos.push_back(_T("I10"));

	vec_Pos.push_back(_T("I17"));
	vec_Pos.push_back(_T("I18"));
	vec_Pos.push_back(_T("I19"));


	vec_Pos.push_back(_T("B11"));
	vec_Pos.push_back(_T("E11"));
	vec_Pos.push_back(_T("H11"));

	vec_Pos.push_back(_T("B20"));
	vec_Pos.push_back(_T("E20"));
	vec_Pos.push_back(_T("H20"));


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("��������������������!"));

		goto GameOver9;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);

GameOver9:
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




void I_2U::OnBnClickedButton8()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

		
	if(m_click==TRUE)
	{
	   m_click=FALSE;
	   AfxBeginThread((AFX_THREADPROC)I_2U::Button8Thread,this,THREAD_PRIORITY_IDLE);
	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}


}



DWORD I_2U::Button8Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

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

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)8));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_tab9);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto GameOver10;

		return TRUE;
	}

	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	vec_float=pMain->CStringtoFloat(pMain->m_tab10);//
	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	
	vec_float=pMain->CStringtoFloat(pMain->m_tab11);//
	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.4f"),vec_float[i]);
		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ
	
	
	for(size_t i=0;i<17;i++)
	{
		vec_str.push_back(_T("��"));
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

	vec_Pos.push_back(_T("I22"));
	vec_Pos.push_back(_T("I23"));
	vec_Pos.push_back(_T("I24"));
	vec_Pos.push_back(_T("I25"));

	vec_Pos.push_back(_T("C32"));
	vec_Pos.push_back(_T("I32"));

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

	vec_Pos.push_back(_T("J22"));
	vec_Pos.push_back(_T("J23"));
	vec_Pos.push_back(_T("J24"));
	vec_Pos.push_back(_T("J25"));


	vec_Pos.push_back(_T("J32"));

	vec_Pos.push_back(_T("B16"));
	vec_Pos.push_back(_T("E16"));
	vec_Pos.push_back(_T("H16"));

	vec_Pos.push_back(_T("B26"));
	vec_Pos.push_back(_T("E26"));
	vec_Pos.push_back(_T("H26"));

	vec_Pos.push_back(_T("B33"));
	vec_Pos.push_back(_T("E33"));
	vec_Pos.push_back(_T("H33"));


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("��������������������!"));

		goto GameOver10;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);


GameOver10:
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


void I_2U::OnBnClickedButton9()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	if(m_click==TRUE)
	{
	   m_click=FALSE;
	   AfxBeginThread((AFX_THREADPROC)I_2U::Button9Thread,this,THREAD_PRIORITY_IDLE);
	}
	else
	{
		AfxMessageBox(_T("�ļ�δ�رգ���������������ȹر��ļ���"));
	}
}



DWORD I_2U::Button9Thread(LPVOID lpParam)
{

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );


	I_2U *pMain=(I_2U*) lpParam;	

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

	pMain->sheet = pMain->sheets.get_Item(COleVariant((short)10));

	CString str_temp; 
	vec_float=pMain->CStringtoFloat(pMain->m_tab12);//
	
	if(vec_float.empty())
	{
		AfxMessageBox(_T("���벻��Ϊ�գ�����!!"));

		goto GameOver11;

		return TRUE;
	}

	for(size_t i=0;i<vec_float.size();i++)
	{

		str_temp.Format(_T("%.3f"),vec_float[i]);
		if(i<2 || (i>7 && i<=9) ||i>=16)
		{
			str_temp+=_T("V");
		}
		else
		{
			str_temp+=_T("mV");
		}

		vec_str.push_back(str_temp);
	    str_temp=_T("");
	}//����������ȡ

	
	
	for(size_t i=0;i<19;i++)
	{
		vec_str.push_back(_T("��"));
	}

	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);
	vec_str.push_back(pMain->m_opname);
	vec_str.push_back(pMain->m_strtime);
	vec_str.push_back(pMain->m_OKorNot);




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


	vec_Pos.push_back(_T("H26"));

	vec_Pos.push_back(_T("C22"));
	vec_Pos.push_back(_T("F22"));
	vec_Pos.push_back(_T("I22"));

	vec_Pos.push_back(_T("C27"));
	vec_Pos.push_back(_T("F27"));
	vec_Pos.push_back(_T("I27"));


	if(vec_Pos.size()!=vec_str.size())
	{
		AfxMessageBox(_T("��������������������!"));

		goto GameOver11;

		return TRUE;
	}

	for(size_t i=0;i<vec_Pos.size();++i)
	{
		pMain->lpDisp = pMain->sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    pMain->range.AttachDispatch(pMain->lpDisp);
	    pMain->range.put_Value(vtMissing, COleVariant(vec_str[i]));
	}

	pMain->app.put_Visible(TRUE);
	pMain->book.Save();
	pMain->book.put_Saved(TRUE);


GameOver11:
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


HBRUSH I_2U::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  �ڴ˸��� DC ���κ�����

	// TODO:  ���Ĭ�ϵĲ������軭�ʣ��򷵻���һ������
	 if(pWnd->GetDlgCtrlID()== IDC_tab2||
		 pWnd->GetDlgCtrlID()== IDC_tab3||
		 pWnd->GetDlgCtrlID()== IDC_tab4||
		 pWnd->GetDlgCtrlID()== IDC_tab5||
		 pWnd->GetDlgCtrlID()== IDC_tab6||
		 pWnd->GetDlgCtrlID()== IDC_tab7||
		 pWnd->GetDlgCtrlID()== IDC_tab8||
		 pWnd->GetDlgCtrlID()== IDC_tab9||
		 pWnd->GetDlgCtrlID()== IDC_tab10||
		 pWnd->GetDlgCtrlID()== IDC_tab11||
		 pWnd->GetDlgCtrlID()== IDC_tab12)
	 {
		 pDC->SetTextColor(RGB(0,0,0)); //������ɫ  
		 pDC->SetBkColor(RGB(53,203,60));
		 pDC->SetBkMode(TRANSPARENT);
		 return (HBRUSH) m_Brush.GetSafeHandle();
	 }


	return hbr;
}


void I_2U::OnBnClickedCancel()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������


	m_closeornot=TRUE;

	if(AfxMessageBox(_T("��ȷ��Ҫ�˳���ǰI-2U�������Ȱ����˼����!"),MB_YESNO|MB_ICONEXCLAMATION)==IDYES)
	{
		CDialogEx::OnCancel();
	}
	
}


void I_2U::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	m_closeornot=TRUE;
	CDialogEx::OnOK();
}

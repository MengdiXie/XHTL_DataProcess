
// inputDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "input.h"
#include "inputDlg.h"
#include "afxdialogex.h"

#include <atlconv.h>
#include <iostream>
#include  <ostream>
#include<algorithm>
#include<io.h>
#include<stdio.h>
#include  "PageOne.h"
#include "I_2U.h"
#include "DO.h"
#include "XHTL.h"
#include "XHTL_ALL.h"



PageOne pageonedlg;
I_2U m_I_2Udlg;
CString exfilename;
CString m_inputopname;
BOOL m_closeornot;
int m_BoardType;
I_2U * m_i_2udlg;
PageOne* m_u_idlg;
DO* m_DOdlg;
XHTL* m_XHTLdlg;
XHTL_ALL* m_XHTL_ALLdlg;
CString m_externinputfilename;
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

 
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CinputDlg �Ի���




CinputDlg::CinputDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CinputDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CinputDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Text(pDX, IDC_EDIT1, m_inputstring1); 
	DDX_Text(pDX, IDC_EDIT2, m_inputstring2);
	DDX_Text(pDX, IDC_EDIT3, m_inputstring3);
	DDX_Text(pDX, IDC_EDIT4, m_inputstring4);
	DDX_Control(pDX, IDC_EDIT6, m_edit);
	DDX_Control(pDX, IDC_WAVE_DRAW, m_picDraw);
	DDX_Control(pDX,IDC_Select,m_Select);
}

BEGIN_MESSAGE_MAP(CinputDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CinputDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CinputDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_colse, &CinputDlg::OnBnClickedcolse)
	ON_BN_CLICKED(IDC_Start, &CinputDlg::OnBnClickedStart)
	ON_BN_CLICKED(IDC_Over, &CinputDlg::OnBnClickedOver)
	ON_WM_CTLCOLOR()
END_MESSAGE_MAP()


// CinputDlg ��Ϣ�������

BOOL CinputDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	m_filenum=0;
	m_inputstring3=_T("��������ǰ��һ�β���");
	m_inputstring4.Format(_T("%d"),m_filenum);
	

	m_Select.AddString(_T("DO"));
	m_Select.AddString(_T("U-I"));
	m_Select.AddString(_T("I-2U"));
	m_Select.AddString(_T("�źŵ�����"));

	
	
	m_Select.SetCurSel(0);

	
	m_i_2udlg=NULL;
	m_closeornot=TRUE;
	UpdateData(FALSE);
	
	strFolder=_T("C:\\�źŵ������\\�������");
	m_getdata=TRUE;
	m_saveclick=FALSE;
	 my_Font.CreatePointFont(140, L"Times NewRoman");//��������ʹ�С

	 m_Brush.CreateSolidBrush(RGB(53,203,60));

	 GetDlgItem(IDC_EDIT1)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT2)->SetFont(&my_Font);
	 GetDlgItem(IDC_EDIT3)->SetFont(&my_Font);

	 m_countclick=0;//��ʼ��
	 m_pictureok=FALSE;

	

	m_BoardType=1;//Ĭ��U-I�忨
	InitEdit();

	 UpdateData(TRUE);
	 m_externinputfilename=m_inputstring3;

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CinputDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CinputDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);



		 


	}
	else
	{
		CDialogEx::OnPaint();
	}

	if(m_pictureok==TRUE)//��ͼ��ʼ���������ԵĻ���ͼ��
	{
	   CRect rectPicture; 
	   m_picDraw.GetClientRect(&rectPicture); 
	   DrawWave(m_picDraw.GetDC(), rectPicture,m_tempv1,m_tempv2,m_tempb1,m_tempb0);
	}

}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CinputDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

std::vector<float> CinputDlg::CStringtoFloat(CString m_s)
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

void CinputDlg::getData(char * filename,std::vector<float> v1,std::vector<float> v2)
{
    //char* path = "C:\\1.txt";
	FILE* fd;
	fd=fopen(filename,"w+");
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
		fclose(fd);
	}

	
	}

}
#define rank 1
void CinputDlg::CalData(FILE * fileID,std::vector<float> v1,std::vector<float> v2)
{
	
	for(size_t i=0;i<v2.size();i++)
		m_out.push_back(v2[i]);


	//double* x=(double*)malloc(sizeof(double)*v1.size());
	//double* y=(double*)malloc(sizeof(double)*v2.size());
	CString strtemp;
	CString strnum;
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

	//���¸�˹������Ԫ��
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
	EditShowData(strtemp);

	for(i=0;i<v1.size();i++)
	{
		fprintf(fileID,"%s%d=%f,detaY=%f\r\n","Fitted Value NO.",i,v1[i]*b[1]+b[0],v2[i]-(v1[i]*b[1]+b[0]));
		strtemp.Format(_T("%s%d=%f,detaY=%f\r\n"),_T("Fitted Value NO."),i,v1[i]*b[1]+b[0],v2[i]-(v1[i]*b[1]+b[0]));
	    EditShowData(strtemp);

		m_out.push_back(v1[i]*b[1]+b[0]);
		m_out.push_back(v2[i]-(v1[i]*b[1]+b[0]));

	}

	fprintf(fileID,"%s\r\n","-----------------------------------");
	fprintf(fileID,"Fitted curve function:y=%f%+fx%\n\n",b[0],b[1]);
	strtemp.Format(_T("%s\r\n"),_T("-----------------------------------"));
    EditShowData(strtemp);
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
    EditShowData(strtemp);

	m_str.push_back(outstr);//����Ϲ�ʽѹ�뻺��

	std::vector<double> detay;

	for(i=0;i<v1.size();i++)
	{
		detay.push_back(abs(b[1]*v1[i]+b[0]-v2[i])/sqrt(b[1]*b[1]+b[0]*b[0]));
	}

	fprintf(fileID,"Linearity=%f\r\n",(*max_element(detay.begin(),detay.end()))/(*max_element(v2.begin(),v2.end())));//�������Զ�
	fprintf(fileID,"%s\r\n","-----------------------------------<<<<");
	strtemp.Format(_T("    Linearity=%f\r\n"),(*max_element(detay.begin(),detay.end()))/(*max_element(v2.begin(),v2.end())));//�������Զ�
    EditShowData(strtemp);
	strtemp.Format(_T("%s\r\n"),_T("----------------------------------------------------"));
    EditShowData(strtemp);
	m_out.push_back((*max_element(detay.begin(),detay.end()))/(*max_element(v2.begin(),v2.end())));//�������Զ� 
	
    CRect rectPicture; 
	m_picDraw.GetClientRect(&rectPicture); 
	DrawWave(m_picDraw.GetDC(), rectPicture,v1,v2,b[1],b[0]);
	m_getdata=FALSE;


	m_tempv1=v1;
	m_tempv2=v2;
	m_tempb1=b[1];
	m_tempb0=b[0];

	m_pictureok=TRUE;



	


}

void CinputDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	static int press=0;
	if(press++==0)//�ж�����ǵ�һ�ΰ��ΰ�ť����������Ч�ļ�·���ȣ�����ǵڶ������ϴ���������ִ�д˼��ģ�飬ȷ�����ļ�����û����Ч�ļ�ʱ������
	{
		OnBnClickedCancel();
	}
	m_countclick++;//ÿ��һ�Σ�������1��������ʱ�������������,�ṩ������ģ���жϣ������ν���һ�δ洢
	EditShowData(CString(_T("")));
	m_getdata=TRUE;


	USES_CONVERSION;
	UpdateData(TRUE);
	CString temp;
	CString outtemp;
	CString strfolder;
	std::vector<float> v_tmp1,v_tmp2;
	v_tmp1=CStringtoFloat(m_inputstring1);
	v_tmp2=CStringtoFloat(m_inputstring2);
	if(v_tmp1.empty() || v_tmp2.empty())
	{
		AfxMessageBox(_T("�����쳣������������!!"));
		return;
	}
	if(v_tmp1.size()!=v_tmp2.size())
	{
		AfxMessageBox(_T("����x��y��������ȣ�����!!"));
		return;
	}
	if(m_filenum<10)
        temp.Format(_T("_SourceData_0%d.txt"),m_filenum++);
	else
	    temp.Format(_T("_SourceData_%2d.txt"),m_filenum++);
	strfolder=strFolder+CString(_T("\\"))+m_inputstring3+temp;//��ʶ
	memcpy(m_filename,(char*)T2A(strfolder.GetBuffer(0)),100);
	
	getData(m_filename,v_tmp1,v_tmp2);
	outtemp=strFolder+CString(_T("\\"))+m_inputstring3+CString(_T("_out.txt"));
	memcpy(m_outfilename,(char*)T2A(outtemp.GetBuffer(0)),100);
	outFileID=fopen(m_outfilename,"a+");
	CalData(outFileID,v_tmp1,v_tmp2);

	fclose(outFileID);


	//���������Ҵ洢
	
	m_inputstring4.Format(_T("%d"),m_filenum);
	UpdateData(FALSE);
	//v_tmp2=CStringtoFloat(m_inputstring2);


	//TRACE("%s\n",m_inputstring1.GetBuffer());

	//TRACE("%s\n",m_inputstring2.GetBuffer());

	//CDialogEx::OnOK();

	





}

void CinputDlg::OnBnClickedCancel()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	USES_CONVERSION;
	UpdateData(TRUE);
	SYSTEMTIME sys; 
    GetLocalTime( &sys ); 
	CString strTemp2;
	CString strTemp1("C:\\�źŵ������\\");
	m_filenum=0;//����
	strFolder=strTemp1+m_inputstring3;

	strTemp2=strFolder+CString(_T("\\*.txt"));
	memcpy(m_Rfilename,(char*)T2A(strTemp2.GetBuffer(0)),100);
	if(!PathFileExists(strFolder))//�ļ��в������򴴽�
    {
        CreateDirectory(strFolder,NULL);
    }
	else
	{
		//���ھ�ɨ�赱ǰ�ļ�����
		long Handle;
		struct _finddata_t FileInfo;
		if((Handle=_findfirst(m_Rfilename,&FileInfo))==-1L)
		{
			//AfxMessageBox(_T("û���ҵ�ƥ�����Ŀ!!"));
			Sleep(1);
		}
		else
		{
			while(_findnext(Handle,&FileInfo)==0)
				m_filenum++;
			_findclose(Handle);
		}
	}

}

bool CinputDlg::InitEdit()
{
	CWnd *pWnd;
	pWnd = GetDlgItem( IDC_EDIT6 );    //��ȡ�ؼ�ָ�룬IDC_EDIT6Ϊ�ؼ�ID��
	//pWnd->MoveWindow( CRect(600,0,1000,1000) );    //�ڴ������Ͻ���ʾһ����100����100�ı༭�ؼ�
	return true;


}
void CinputDlg::EditShowData(CString _s)
{
	CString cstr_temp;
	CString cstr_temp02;
	if(m_getdata)
	{
	   m_edit.GetWindowText(cstr_temp);
	   cstr_temp02=cstr_temp+_s;
	}
	else
	{
		cstr_temp02=CStringA(_T(""));
	}
	m_edit.SetWindowText(cstr_temp02);
}
void CinputDlg::DrawWave(CDC *pDC,CRect &rectPicture,std::vector<float> x,std::vector<float> y,float a,float b)
{
	float fDeltaX;     // x������������ͼ����������   
    float fDeltaY;     // y��ÿ���߼���λ��Ӧ������ֵ   
    int nX;      // ������ʱ���ڴ洢��ͼ��ĺ�����   
    int nY;      // ������ʱ���ڴ洢��ͼ��������� 

	CPen newPen;       // ���ڴ����»���   
    CPen *pOldPen;     // ���ڴ�žɻ���   
	CPen redpen;
    CBrush newBrush;   // ���ڴ����»�ˢ   
    CBrush *pOldBrush; // ���ڴ�žɻ�ˢ   


	float xmax=*max_element(x.begin(),x.end());
	float ymax=*max_element(y.begin(),y.end());
	// ����fDeltaX��fDeltaY   
    fDeltaX = (float)rectPicture.Width() /(xmax+1);   
    fDeltaY = (float)rectPicture.Height() /(ymax+1);   

	// ������ɫ�»�ˢ   
    newBrush.CreateSolidBrush(RGB(0,0,0));  

	// ѡ���»�ˢ�������ɻ�ˢ��ָ�뱣�浽pOldBrush   
    pOldBrush = pDC->SelectObject(&newBrush);   
    // �Ժ�ɫ��ˢΪ��ͼ�ؼ�����ɫ���γɺ�ɫ����   
    pDC->Rectangle(rectPicture);   


	 // �ָ��ɻ�ˢ   
    pDC->SelectObject(pOldBrush);   
    // ɾ���»�ˢ   
    newBrush.DeleteObject();   


	// ����ʵ�Ļ��ʣ��ֶ�Ϊ1����ɫΪ��ɫ   
    newPen.CreatePen(PS_SOLID, 1, RGB(0,255,0));   
    // ѡ���»��ʣ������ɻ��ʵ�ָ�뱣�浽pOldPen   
    pOldPen = pDC->SelectObject(&newPen);   
  
    // ����ǰ���ƶ�����ͼ�ؼ����ڵ����½ǣ��Դ�Ϊ���ε���ʼ��   
    pDC->MoveTo(rectPicture.left, rectPicture.bottom);   

	//�ƶ�
	nX=rectPicture.left + (int)(x[0]* fDeltaX);  
	nY = rectPicture.bottom - (int)((a*x[0]+b)* fDeltaY);  
	pDC->MoveTo(nX, nY);

	nX=rectPicture.left + (int)(x[x.size()-1]* fDeltaX);  
	nY = rectPicture.bottom - (int)((a*x[x.size()-1]+b)* fDeltaY);  

	 pDC->LineTo(nX, nY);   
 
	// ����m_nzValues������ÿ�����Ӧ������λ�ã����������ӣ������γ�����   
	CPoint Point1,Point2;
	redpen.CreatePen(PS_SOLID, 2, RGB(255,0,0));   
    
    pDC->SelectObject(&redpen);   
    for (int i=0; i<x.size(); i++)   
    {   
        nX = rectPicture.left + (int)(x[i] * fDeltaX);   
        nY = rectPicture.bottom - (int)(y[i] * fDeltaY);   
        //pDC->LineTo(nX, nY);   

		//pDC.dot(CPoint(nX,nY),RGB(255,0,0));
		pDC->SetPixel(nX,nY,RGB(255,0,0));

		pDC->SetPixel(nX-2, nY, RGB(255,0,0));
        pDC->SetPixel(nX+2, nY, RGB(255,0,0));
        pDC->SetPixel(nX, nY, RGB(255,0,0));
        pDC->SetPixel(nX, nY-2, RGB(255,0,0));
        pDC->SetPixel(nX, nY+2, RGB(255,0,0));


       
	   pDC->MoveTo(nX-2, nY); 
	   pDC->LineTo(nX+2, nY);  
	   pDC->MoveTo(nX, nY-2); 
	   pDC->LineTo(nX, nY+2); 

#if 0
		Point1.x=nX-3;
		Point1.y=nY-3;
		Point2.x=nX+3;
		Point2.y=nY+3;

		pDC->Ellipse(CRect(Point1,Point2));
#endif


    }   

	CString str_temp;
	

   	if((int)b>=0 && (int)a>=0)
		str_temp.Format(_T(" f(x)=%.4f+%.4fx "),b,a);
	else if((int)b>=0 && a<0)
		str_temp.Format(_T(" f(x)=%.4f%.4fx "),b,a);
	else if(b<0 && a<0)
		str_temp.Format(_T(" f(x)=%.4f%.4fx "),b,a);
	else if(b<0 && (int)a>=0)
		str_temp.Format(_T(" f(x)=%.4f+%.4fx "),b,a);

	pDC->TextOutW(rectPicture.left+rectPicture.Width()/6,rectPicture.bottom-6*rectPicture.Height()/7,str_temp);

	// �ָ��ɻ���   
    pDC->SelectObject(pOldPen);   
    // ɾ���»���   
    newPen.DeleteObject();   
}


void CinputDlg::OnBnClickedcolse()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������



	if(AfxMessageBox(_T("��ȷ��Ҫ�˳���ǰ����¼�빤��������˼����!"),MB_YESNO|MB_ICONEXCLAMATION)==IDYES)
	{
		CinputDlg::OnCancel();
	}

	

	
}


void CinputDlg::OnBnClickedStart()
{
	if(m_saveclick==FALSE)
	{
		m_saveclick=TRUE;
	}
	else
	{
		AfxMessageBox(_T("���Ѿ���ǰ���У�����ȷ���Ѿ�����������������"));
		return;
	}


	// TODO: �ڴ���ӿؼ�֪ͨ����������
	GetDlgItem(IDCANCEL)->EnableWindow (FALSE);//��ť������
	UpdateData(TRUE);
	CString strTemp1("C:\\�źŵ������\\");
	CString exstrFolder;
	exstrFolder=strTemp1+m_inputstring3;
	int nIndex=m_Select.GetCurSel();
	if(!PathFileExists(exstrFolder))//������ʾ�����µ��ļ������ڴ�Ŵ˴β��Եı��
    {
        CreateDirectory(exstrFolder,NULL);


		//�����ļ�
	    if(nIndex==UI)
		{
	        exstrFolder=strTemp1+CString("sample_U_I.xlsx");//ȷ�����ļ��洢·����ͨ�������ļ��ķ�ʽ�����Ʊ�񸱱�
			m_inputstring1=_T("0,1,2,3,4,5");
		}
		else if(nIndex==I2U)
		{
			 exstrFolder=strTemp1+CString("sample_I_2U.xlsx");//ȷ�����ļ��洢·����ͨ�������ļ��ķ�ʽ�����Ʊ�񸱱�
			 m_inputstring1=_T("4,8,12,16,20");
		}
		else if(nIndex==D_O)
		{
			exstrFolder=strTemp1+CString("sample_DO.xlsx");
            GetDlgItem(IDOK)->EnableWindow (FALSE);//��ť������


		}
		else if(nIndex==XHTL_ID)
		{
			exstrFolder=strTemp1+CString("sample_XHTL_ALL.xlsx");
            GetDlgItem(IDOK)->EnableWindow (FALSE);//��ť������


		}


	    exfilename=strTemp1+m_inputstring3+CString("\\")+m_inputstring3+CString(".xlsx");
	   //���Ʋ���
	   CopyFile(exstrFolder,exfilename,FALSE);
    }

	else{

               if(nIndex==UI)
		        {
	              exstrFolder=strTemp1+CString("sample_U_I.xlsx");//ȷ�����ļ��洢·����ͨ�������ļ��ķ�ʽ�����Ʊ�񸱱�
				  m_inputstring1=_T("0,1,2,3,4,5");
		        }
		        else if(nIndex==I2U)
		       {
			       exstrFolder=strTemp1+CString("sample_I_2U.xlsx");//ȷ�����ļ��洢·����ͨ�������ļ��ķ�ʽ�����Ʊ�񸱱�
				   m_inputstring1=_T("4,8,12,16,20");
			   }
			   	else if(nIndex==D_O)
		       {
			       exstrFolder=strTemp1+CString("sample_DO.xlsx");
                   GetDlgItem(IDOK)->EnableWindow (FALSE);//��ť������
                }
				else if(nIndex==XHTL_ID)
		        {
			       exstrFolder=strTemp1+CString("sample_XHTL_ALL.xlsx");
                   GetDlgItem(IDOK)->EnableWindow (FALSE);//��ť������
		}




		       exfilename=strTemp1+m_inputstring3+CString("\\")+m_inputstring3+CString(".xlsx");
			   int nRet = _taccess(exfilename, 0);
               int ret_out;
			   ret_out=( 0 == nRet || EACCES == nRet);
			   if(ret_out==1)
			   {
			   }
			   else if(ret_out!=1)
			   {
				   CopyFile(exstrFolder,exfilename,FALSE);//�ļ��д��ں�Ͳ����ظ���������Ҫ
			   }
	}

	UpdateData(FALSE);


	
#if 0	
	pageonedlg.app=app;
    pageonedlg.books=books;
    pageonedlg.book=book;
    pageonedlg.sheets=sheets;
    pageonedlg.sheet=sheet;
    pageonedlg.range=range;
    pageonedlg.iCell=iCell;
    pageonedlg.lpDisp=lpDisp;
#endif

	
	m_BoardType=nIndex;
	if(nIndex==UI)//ѡ��U-I�忨
	{
		
		if(m_u_idlg!=NULL)
		{
			delete m_u_idlg;
		}
		m_u_idlg =new PageOne;


		m_u_idlg->Create(IDD_SHEET1,GetDesktopWindow());
	    m_u_idlg->ShowWindow(SW_SHOW);


	    //if(m_u_idlg->DoModal()==IDOK)
	    {


	     }
	}
	else if(nIndex==I2U)//ѡ��I-2U�忨
	{
	    //if(m_I_2Udlg.DoModal()==IDOK)
	    {

			if(m_i_2udlg!=NULL)//���������������
			{
				delete m_i_2udlg;
			}
			
			m_i_2udlg=new I_2U;

			m_i_2udlg->Create(IDD_I_2U,GetDesktopWindow());
		    m_i_2udlg->ShowWindow(SW_SHOW);

#if 0
		     	if(m_I_2Udlg.m_hWnd!=NULL)
				{
					m_I_2Udlg.DestroyWindow();
				}

		    	m_I_2Udlg.Create(IDD_I_2U,GetDesktopWindow());
			    m_I_2Udlg.ShowWindow(SW_SHOW);
#endif
	
	     }
	}
	else if(nIndex==D_O)//ѡ��DO�忨
	{
		if(m_DOdlg!=NULL)
		{
			delete m_DOdlg;
		}
		m_DOdlg =new DO;

	
	    m_DOdlg->Create(IDD_DO,GetDesktopWindow());
		m_DOdlg->ShowWindow(SW_SHOW);
	}
	else if(nIndex==XHTL_ID)
	{
		m_externinputfilename=m_inputstring3;
		if(m_XHTL_ALLdlg!=NULL)
		{
			delete m_XHTL_ALLdlg;
		}
		m_XHTL_ALLdlg=new XHTL_ALL;
		m_XHTL_ALLdlg->Create(IDD_XHTLALL,GetDesktopWindow());
		m_XHTL_ALLdlg->ShowWindow(SW_SHOW);
	}



}


void CinputDlg::OnBnClickedOver()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_saveclick==TRUE)
	{
		
	}
	else
	{
		AfxMessageBox(_T("Ŀǰû�������������ȵ������ʼ�������"));
		return;
	}


	int nIndex=m_Select.GetCurSel();

	if(nIndex==D_O)
	{
		 GetDlgItem(IDOK)->EnableWindow (TRUE);//��ť������
		 GetDlgItem(IDCANCEL)->EnableWindow (TRUE);//��ť������
		 return;
	}

	if(nIndex==XHTL_ID)
	{

		GetDlgItem(IDOK)->EnableWindow (TRUE);//��ť������
	    GetDlgItem(IDCANCEL)->EnableWindow (TRUE);//��ť������
		m_saveclick=FALSE;

		return;

	}





	if(m_countclick<6)
	{
	    //m_out.clear();
	    //m_str.clear();//����
		//m_countclick=0;
		AfxMessageBox(L"�������ݲ���6�飬����,�����������!");
		return;
	}
	m_countclick-=6;

	HRESULT hr = ::CoInitializeEx( NULL, COINIT_MULTITHREADED );
	std::vector<CString> vec_str;
	std::vector<float> vec_float;
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


	
	if(nIndex==UI)//U-I
	{


	sheet = sheets.get_Item(COleVariant((short)5));
	CString str_temp; 
	for(size_t i=0;i<m_out.size();i++)
	{

		str_temp.Format(_T("%.4f"),m_out[i]);
		if(i==18)
			str_temp.Format(_T("%.6f"),m_out[i]);
		if(i==38)
			str_temp.Format(_T("%.6f"),m_out[i]);
		if(i==56)
			str_temp.Format(_T("%.6f"),m_out[i]);
		vec_str.push_back(str_temp);
		str_temp=_T("");
		if(i==56)
			break;
	}

		////////
    for(size_t i=0;i<18;i++)
     	vec_str.push_back(_T("��"));


	for(size_t i=0;i<m_str.size();i++)
	{

		vec_str.push_back(m_str[i]);
		if(i==2)
		   break;
	}

	
	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//��ӡ�������

	vec_str.push_back(CString(m_inputopname));
	vec_str.push_back(strtime);
	vec_str.push_back(_T("�ϸ�� ���ϸ�"));



	vec_Pos.push_back(_T("I11"));
	vec_Pos.push_back(_T("I12"));
	vec_Pos.push_back(_T("I13"));
	vec_Pos.push_back(_T("I14"));
	vec_Pos.push_back(_T("I15"));
	vec_Pos.push_back(_T("I16"));
	vec_Pos.push_back(_T("J11"));
	vec_Pos.push_back(_T("K11"));
	vec_Pos.push_back(_T("J12"));
	vec_Pos.push_back(_T("K12"));
	vec_Pos.push_back(_T("J13"));
	vec_Pos.push_back(_T("K13"));
	vec_Pos.push_back(_T("J14"));
	vec_Pos.push_back(_T("K14"));
	vec_Pos.push_back(_T("J15"));
	vec_Pos.push_back(_T("K15"));
	vec_Pos.push_back(_T("J16"));
	vec_Pos.push_back(_T("K16"));
	vec_Pos.push_back(_T("K17"));

	vec_Pos.push_back(_T("I18"));
	vec_Pos.push_back(_T("I19"));
	vec_Pos.push_back(_T("I20"));
	vec_Pos.push_back(_T("I21"));
	vec_Pos.push_back(_T("I22"));
	vec_Pos.push_back(_T("I23"));
	vec_Pos.push_back(_T("J18"));
	vec_Pos.push_back(_T("K18"));
	vec_Pos.push_back(_T("J19"));
	vec_Pos.push_back(_T("K19"));
	vec_Pos.push_back(_T("J20"));
	vec_Pos.push_back(_T("K20"));
	vec_Pos.push_back(_T("J21"));
	vec_Pos.push_back(_T("K21"));
	vec_Pos.push_back(_T("J22"));
	vec_Pos.push_back(_T("K22"));
	vec_Pos.push_back(_T("J23"));
	vec_Pos.push_back(_T("K23"));
	vec_Pos.push_back(_T("K24"));

	vec_Pos.push_back(_T("I25"));
	vec_Pos.push_back(_T("I26"));
	vec_Pos.push_back(_T("I27"));
	vec_Pos.push_back(_T("I28"));
	vec_Pos.push_back(_T("I29"));
	vec_Pos.push_back(_T("I30"));
	
	vec_Pos.push_back(_T("J25"));
	vec_Pos.push_back(_T("K25"));
	vec_Pos.push_back(_T("J26"));
	vec_Pos.push_back(_T("K26"));
	vec_Pos.push_back(_T("J27"));
	vec_Pos.push_back(_T("K27"));
	vec_Pos.push_back(_T("J28"));
	vec_Pos.push_back(_T("K28"));
	vec_Pos.push_back(_T("J29"));
	vec_Pos.push_back(_T("K29"));
	vec_Pos.push_back(_T("J30"));
	vec_Pos.push_back(_T("K30"));
	vec_Pos.push_back(_T("K31"));



	

	
	vec_Pos.push_back(_T("L11"));
	vec_Pos.push_back(_T("L12"));
	vec_Pos.push_back(_T("L13"));
	vec_Pos.push_back(_T("L14"));
	vec_Pos.push_back(_T("L15"));
	vec_Pos.push_back(_T("L16"));

	vec_Pos.push_back(_T("L18"));
	vec_Pos.push_back(_T("L19"));
	vec_Pos.push_back(_T("L20"));
	vec_Pos.push_back(_T("L21"));
	vec_Pos.push_back(_T("L22"));
	vec_Pos.push_back(_T("L23"));
	
	vec_Pos.push_back(_T("L25"));
	vec_Pos.push_back(_T("L26"));
	vec_Pos.push_back(_T("L27"));
	vec_Pos.push_back(_T("L28"));
	vec_Pos.push_back(_T("L29"));
	vec_Pos.push_back(_T("L30"));


	vec_Pos.push_back(_T("E17"));
	vec_Pos.push_back(_T("E24"));
	vec_Pos.push_back(_T("E31"));

	vec_Pos.push_back(_T("D32"));
	vec_Pos.push_back(_T("H32"));
	vec_Pos.push_back(_T("K32"));


	if(vec_Pos.size()!=vec_str.size())
	{
		MessageBox(_T("U-I ��������������������!"));

		goto GameOver2;
		return;
	}
	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));
		
	}

	sheet = sheets.get_Item(COleVariant((short)6));

	vec_str.clear();
	vec_Pos.clear();
	if(m_out.size()>57)
	{
	   for(size_t i=57;i<m_out.size();i++)
	   {

		str_temp.Format(_T("%.4f"),m_out[i]);
		if(i==(18+57))
			str_temp.Format(_T("%.6f"),m_out[i]);
		if(i==(38+57))
			str_temp.Format(_T("%.6f"),m_out[i]);
		if(i==(56+57))
			str_temp.Format(_T("%.6f"),m_out[i]);
		vec_str.push_back(str_temp);
		str_temp=_T("");
		if(i==(56+57))
			break;
	    }
	}

		////////
    for(size_t i=0;i<18;i++)
     	vec_str.push_back(_T("��"));


	for(size_t i=3;i<m_str.size();i++)
	{
		vec_str.push_back(m_str[i]);
		if(i==(2+3))
		   break;
	}

	vec_str.push_back(CString(m_inputopname));
	vec_str.push_back(strtime);
	vec_str.push_back(_T("�ϸ�� ���ϸ�"));

	vec_Pos.push_back(_T("I18"));
	vec_Pos.push_back(_T("I19"));
	vec_Pos.push_back(_T("I20"));
	vec_Pos.push_back(_T("I21"));
	vec_Pos.push_back(_T("I22"));
	vec_Pos.push_back(_T("I23"));
	vec_Pos.push_back(_T("J18"));
	vec_Pos.push_back(_T("K18"));
	vec_Pos.push_back(_T("J19"));
	vec_Pos.push_back(_T("K19"));
	vec_Pos.push_back(_T("J20"));
	vec_Pos.push_back(_T("K20"));
	vec_Pos.push_back(_T("J21"));
	vec_Pos.push_back(_T("K21"));
	vec_Pos.push_back(_T("J22"));
	vec_Pos.push_back(_T("K22"));
	vec_Pos.push_back(_T("J23"));
	vec_Pos.push_back(_T("K23"));
	vec_Pos.push_back(_T("K24"));

	vec_Pos.push_back(_T("I25"));
	vec_Pos.push_back(_T("I26"));
	vec_Pos.push_back(_T("I27"));
	vec_Pos.push_back(_T("I28"));
	vec_Pos.push_back(_T("I29"));
	vec_Pos.push_back(_T("I30"));
	vec_Pos.push_back(_T("J25"));
	vec_Pos.push_back(_T("K25"));
	vec_Pos.push_back(_T("J26"));
	vec_Pos.push_back(_T("K26"));
	vec_Pos.push_back(_T("J27"));
	vec_Pos.push_back(_T("K27"));
	vec_Pos.push_back(_T("J28"));
	vec_Pos.push_back(_T("K28"));
	vec_Pos.push_back(_T("J29"));
	vec_Pos.push_back(_T("K29"));
	vec_Pos.push_back(_T("J30"));
	vec_Pos.push_back(_T("K30"));
	vec_Pos.push_back(_T("K31"));

	vec_Pos.push_back(_T("I32"));
	vec_Pos.push_back(_T("I33"));
	vec_Pos.push_back(_T("I34"));
	vec_Pos.push_back(_T("I35"));
	vec_Pos.push_back(_T("I36"));
	vec_Pos.push_back(_T("I37"));
	
	vec_Pos.push_back(_T("J32"));
	vec_Pos.push_back(_T("K32"));
	vec_Pos.push_back(_T("J33"));
	vec_Pos.push_back(_T("K33"));
	vec_Pos.push_back(_T("J34"));
	vec_Pos.push_back(_T("K34"));
	vec_Pos.push_back(_T("J35"));
	vec_Pos.push_back(_T("K35"));
	vec_Pos.push_back(_T("J36"));
	vec_Pos.push_back(_T("K36"));
	vec_Pos.push_back(_T("J37"));
	vec_Pos.push_back(_T("K37"));
	vec_Pos.push_back(_T("K38"));



	

	
	vec_Pos.push_back(_T("L18"));
	vec_Pos.push_back(_T("L19"));
	vec_Pos.push_back(_T("L20"));
	vec_Pos.push_back(_T("L21"));
	vec_Pos.push_back(_T("L22"));
	vec_Pos.push_back(_T("L23"));

	vec_Pos.push_back(_T("L25"));
	vec_Pos.push_back(_T("L26"));
	vec_Pos.push_back(_T("L27"));
	vec_Pos.push_back(_T("L28"));
	vec_Pos.push_back(_T("L29"));
	vec_Pos.push_back(_T("L30"));
	
	vec_Pos.push_back(_T("L32"));
	vec_Pos.push_back(_T("L33"));
	vec_Pos.push_back(_T("L34"));
	vec_Pos.push_back(_T("L35"));
	vec_Pos.push_back(_T("L36"));
	vec_Pos.push_back(_T("L37"));


	vec_Pos.push_back(_T("E24"));
	vec_Pos.push_back(_T("E31"));
	vec_Pos.push_back(_T("E38"));

	vec_Pos.push_back(_T("C39"));
	vec_Pos.push_back(_T("G39"));
	vec_Pos.push_back(_T("K39"));


	if(vec_Pos.size()!=vec_str.size())
	{
		MessageBox(_T("I-2U ��������������������!"));

		goto GameOver2;
		return;
	 }

	}
	else if(nIndex==I2U)//I-2U�忨
	{
	sheet = sheets.get_Item(COleVariant((short)9));
	CString str_temp; 
	for(size_t i=0;i<m_out.size();i++)
	{

		str_temp.Format(_T("%.4f"),m_out[i]);
		if(i==15)
			str_temp.Format(_T("%.6f"),m_out[i]);
		if(i==31)
			str_temp.Format(_T("%.6f"),m_out[i]);
		if(i==47||i==63||i==79 ||i==95)
			str_temp.Format(_T("%.6f"),m_out[i]);
		


		vec_str.push_back(str_temp);
		str_temp=_T("");
		if(i==95)
			break;
	}

	for(size_t i=0;i<30;i++)
     	vec_str.push_back(_T("��"));


	for(size_t i=0;i<m_str.size();i++)
	{

		vec_str.push_back(m_str[i]);
		if(i==5)
		   break;
	}

	COleDateTime t = COleDateTime::GetCurrentTime();

	CString strtime = t.Format(_T("%Y/%m/%d"));//��ӡ�������

	vec_str.push_back(CString(m_inputopname));
	vec_str.push_back(strtime);
	vec_str.push_back(_T("�ϸ�� ���ϸ�"));





	vec_Pos.push_back(_T("I4"));
	vec_Pos.push_back(_T("I5"));
	vec_Pos.push_back(_T("I6"));
	vec_Pos.push_back(_T("I7"));
	vec_Pos.push_back(_T("I8"));
	vec_Pos.push_back(_T("J4"));
	vec_Pos.push_back(_T("K4"));
	vec_Pos.push_back(_T("J5"));
	vec_Pos.push_back(_T("K5"));
	vec_Pos.push_back(_T("J6"));
	vec_Pos.push_back(_T("K6"));
	vec_Pos.push_back(_T("J7"));
	vec_Pos.push_back(_T("K7"));
	vec_Pos.push_back(_T("J8"));
	vec_Pos.push_back(_T("K8"));
	vec_Pos.push_back(_T("K9"));




	vec_Pos.push_back(_T("I10"));
	vec_Pos.push_back(_T("I11"));
	vec_Pos.push_back(_T("I12"));
	vec_Pos.push_back(_T("I13"));
	vec_Pos.push_back(_T("I14"));
	vec_Pos.push_back(_T("J10"));
	vec_Pos.push_back(_T("K10"));
	vec_Pos.push_back(_T("J11"));
	vec_Pos.push_back(_T("K11"));
	vec_Pos.push_back(_T("J12"));
	vec_Pos.push_back(_T("K12"));
	vec_Pos.push_back(_T("J13"));
	vec_Pos.push_back(_T("K13"));
	vec_Pos.push_back(_T("J14"));
	vec_Pos.push_back(_T("K14"));
	vec_Pos.push_back(_T("K15"));


	vec_Pos.push_back(_T("I16"));
	vec_Pos.push_back(_T("I17"));
	vec_Pos.push_back(_T("I18"));
	vec_Pos.push_back(_T("I19"));
	vec_Pos.push_back(_T("I20"));
	vec_Pos.push_back(_T("J16"));
	vec_Pos.push_back(_T("K16"));
	vec_Pos.push_back(_T("J17"));
	vec_Pos.push_back(_T("K17"));
	vec_Pos.push_back(_T("J18"));
	vec_Pos.push_back(_T("K18"));
	vec_Pos.push_back(_T("J19"));
	vec_Pos.push_back(_T("K19"));
	vec_Pos.push_back(_T("J20"));
	vec_Pos.push_back(_T("K20"));
	vec_Pos.push_back(_T("K21"));


	vec_Pos.push_back(_T("I22"));
	vec_Pos.push_back(_T("I23"));
	vec_Pos.push_back(_T("I24"));
	vec_Pos.push_back(_T("I25"));
	vec_Pos.push_back(_T("I26"));
	vec_Pos.push_back(_T("J22"));
	vec_Pos.push_back(_T("K22"));
	vec_Pos.push_back(_T("J23"));
	vec_Pos.push_back(_T("K23"));
	vec_Pos.push_back(_T("J24"));
	vec_Pos.push_back(_T("K24"));
	vec_Pos.push_back(_T("J25"));
	vec_Pos.push_back(_T("K25"));
	vec_Pos.push_back(_T("J26"));
	vec_Pos.push_back(_T("K26"));
	vec_Pos.push_back(_T("K27"));

	vec_Pos.push_back(_T("I28"));
	vec_Pos.push_back(_T("I29"));
	vec_Pos.push_back(_T("I30"));
	vec_Pos.push_back(_T("I31"));
	vec_Pos.push_back(_T("I32"));
	vec_Pos.push_back(_T("J28"));
	vec_Pos.push_back(_T("K28"));
	vec_Pos.push_back(_T("J29"));
	vec_Pos.push_back(_T("K29"));
	vec_Pos.push_back(_T("J30"));
	vec_Pos.push_back(_T("K30"));
	vec_Pos.push_back(_T("J31"));
	vec_Pos.push_back(_T("K31"));
	vec_Pos.push_back(_T("J32"));
	vec_Pos.push_back(_T("K32"));
	vec_Pos.push_back(_T("K33"));


	vec_Pos.push_back(_T("I34"));
	vec_Pos.push_back(_T("I35"));
	vec_Pos.push_back(_T("I36"));
	vec_Pos.push_back(_T("I37"));
	vec_Pos.push_back(_T("I38"));
	vec_Pos.push_back(_T("J34"));
	vec_Pos.push_back(_T("K34"));
	vec_Pos.push_back(_T("J35"));
	vec_Pos.push_back(_T("K35"));
	vec_Pos.push_back(_T("J36"));
	vec_Pos.push_back(_T("K36"));
	vec_Pos.push_back(_T("J37"));
	vec_Pos.push_back(_T("K37"));
	vec_Pos.push_back(_T("J38"));
	vec_Pos.push_back(_T("K38"));
	vec_Pos.push_back(_T("K39"));


	vec_Pos.push_back(_T("L4"));
	vec_Pos.push_back(_T("L5"));
	vec_Pos.push_back(_T("L6"));
	vec_Pos.push_back(_T("L7"));
	vec_Pos.push_back(_T("L8"));

	vec_Pos.push_back(_T("L10"));
	vec_Pos.push_back(_T("L11"));
	vec_Pos.push_back(_T("L12"));
	vec_Pos.push_back(_T("L13"));
	vec_Pos.push_back(_T("L14"));

	vec_Pos.push_back(_T("L16"));
	vec_Pos.push_back(_T("L17"));
	vec_Pos.push_back(_T("L18"));
	vec_Pos.push_back(_T("L19"));
	vec_Pos.push_back(_T("L20"));

	vec_Pos.push_back(_T("L22"));
	vec_Pos.push_back(_T("L23"));
	vec_Pos.push_back(_T("L24"));
	vec_Pos.push_back(_T("L25"));
	vec_Pos.push_back(_T("L26"));


	vec_Pos.push_back(_T("L28"));
	vec_Pos.push_back(_T("L29"));
	vec_Pos.push_back(_T("L30"));
	vec_Pos.push_back(_T("L31"));
	vec_Pos.push_back(_T("L32"));

	vec_Pos.push_back(_T("L34"));
	vec_Pos.push_back(_T("L35"));
	vec_Pos.push_back(_T("L36"));
	vec_Pos.push_back(_T("L37"));
	vec_Pos.push_back(_T("L38"));

	vec_Pos.push_back(_T("E9"));
	vec_Pos.push_back(_T("E15"));
	vec_Pos.push_back(_T("E21"));
	vec_Pos.push_back(_T("E27"));
	vec_Pos.push_back(_T("E33"));
	vec_Pos.push_back(_T("E39"));

	vec_Pos.push_back(_T("D40"));
	vec_Pos.push_back(_T("H40"));
	vec_Pos.push_back(_T("K40"));

	if(vec_Pos.size()!=vec_str.size())
	{
		MessageBox(_T("��������������ʵ����������������������!"));

		goto GameOver2;
		return;
	 }

	}
	for(size_t i=0;i<vec_Pos.size();++i)
	{
		lpDisp = sheet.get_Range(COleVariant(vec_Pos[i]), COleVariant(vec_Pos[i]));
	    //���������ӵ���Ԫ��//������д���Ӧ�ĵ�Ԫ��
	    range.AttachDispatch(lpDisp);
	    range.put_Value(vtMissing, COleVariant(vec_str[i]));	
	}

GameOver2:

	  app.put_Visible(TRUE);
	  book.Save();
	  book.put_Saved(TRUE);

	  m_out.clear();
	  m_str.clear();
	 // m_filenum=0;//�������ʱ��Ҫ����



	  books.Close(); 
      app.Quit();  			// �˳�
	  //�ͷŶ���  
	  range.ReleaseDispatch();
	  sheet.ReleaseDispatch();
	  sheets.ReleaseDispatch();
	  book.ReleaseDispatch();
	  books.ReleaseDispatch();
	  app.ReleaseDispatch();

      GetDlgItem(IDCANCEL)->EnableWindow (TRUE);
	


	  CoUninitialize(); 

	  m_saveclick=FALSE;

}


HBRUSH CinputDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  �ڴ˸��� DC ���κ�����

	 if(pWnd->GetDlgCtrlID()== IDC_EDIT1||
		 pWnd->GetDlgCtrlID()== IDC_EDIT2||
		 pWnd->GetDlgCtrlID()== IDC_EDIT3)

	 {
		 pDC->SetTextColor(RGB(0,0,0)); //������ɫ  
		 pDC->SetBkColor(RGB(53,203,60));
		 pDC->SetBkMode(TRANSPARENT);
		 return (HBRUSH) m_Brush.GetSafeHandle();
	 }


	// TODO:  ���Ĭ�ϵĲ������軭�ʣ��򷵻���һ������
	return hbr;
}

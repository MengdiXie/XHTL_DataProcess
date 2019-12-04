
// inputDlg.h : ͷ�ļ�
//


#pragma once
#include<vector>

#include "CApplication.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"


// CinputDlg �Ի���
class CinputDlg : public CDialogEx
{
// ����
public:
	CinputDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_INPUT_DIALOG };

	enum {D_O=0,I2U=1,UI=2,XHTL_ID=3};

	CString m_inputstring1;
	CString m_inputstring2;
	CString m_inputstring3;
	CString m_inputstring4;
	CString strFolder;
	int m_filenum;
	FILE * outFileID;
	bool m_getdata;
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��

	std::vector<float> CStringtoFloat(CString m_s);

	void getData(char * filename,std::vector<float> v1,std::vector<float> v2);

	void CalData(FILE * fileID,std::vector<float> v1,std::vector<float> v2);

	void EditShowData(CString _s);

	void DrawWave(CDC *pDC,CRect &rectPicture,std::vector<float> x,std::vector<float> y,float a,float b);

	char m_filename[100];
	char m_outfilename[100];
	char m_Rfilename[100];


	bool InitEdit();

	
	CApplication app;
    CWorkbooks books;
    CWorkbook book;
    CWorksheets sheets;
    CWorksheet sheet;
    CRange range;
    CRange iCell;
    LPDISPATCH lpDisp;


	std::vector<float> m_out;
	std::vector<CString> m_str;


	CComboBox m_Select;
	
     CFont my_Font;


	CBrush m_Brush;

	int m_countclick;//��¼��������ʱ�����������У��


	BOOL m_saveclick;


	std::vector<float> m_tempv1;
	std::vector<float> m_tempv2;
	float m_tempb1;
	float m_tempb0;

	BOOL m_pictureok;

// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedcolse();
	CEdit m_edit;
	CStatic m_picDraw;
	afx_msg void OnBnClickedStart();
	afx_msg void OnBnClickedOver();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
};

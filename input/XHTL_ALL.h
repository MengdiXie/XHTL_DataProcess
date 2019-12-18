#pragma once


// XHTL_ALL 对话框
#include"inputDlg.h"
class XHTL_ALL : public CDialogEx
{
	DECLARE_DYNAMIC(XHTL_ALL)

public:
	XHTL_ALL(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~XHTL_ALL();

// 对话框数据
	enum { IDD = IDD_XHTLALL };
	enum {U_I_1=0,U_I_2=1,U_I_3=2,U_I_4=3,I_2U_1=4,I_2U_2=5,I_2U_3=6,I_2U_4=7,I_2U_5=8,I_2U_6=9,I_2U_7=10,I_2U_8=11,I_2U_9=12,I_2U_10=13,I_2U_11=14,I_2U_12=15,I_2U_13=16,I_2U_14=17,I_2U_15=18};
	CApplication app;
    CWorkbooks books;
    CWorkbook book;
    CWorksheets sheets;
    CWorksheet sheet;
    CRange range;
    CRange iCell;
    LPDISPATCH lpDisp;

	CFont my_Font;


	CBrush m_Brush;
	CComboBox m_Select1;
	CComboBox m_Select2;
	BOOL m_checkbox1;
	BOOL m_checkbox2;

	BOOL m_checkbox3;
	BOOL m_checkbox4;

	BOOL m_buttonclick;

	CString m_IDC_EDIT8;
	CString m_IDC_EDIT9;
	CString m_IDC_EDIT14;
	CString m_basename;
	CString strfolder;
	CString m_strtime;

	CString m_IDC_EDIT1;
	CString m_IDC_EDIT2;
	CString m_IDC_EDIT3;
	CString m_IDC_EDIT4;
	CString m_IDC_EDIT5;
	CString m_IDC_EDIT6;
	CString m_IDC_EDIT7;
	CString m_IDC_EDIT10;
	CString m_IDC_EDIT15;
	CString m_IDC_EDIT16;
	CString m_IDC_EDIT17;
	char m_filename[100];
	char m_outfilename[100];
	int m_filenum;
	FILE * outFileID;

	std::vector<float> m_deta;

	std::vector<float> CStringtoFloat(CString m_s);
    void getData(char * filename,std::vector<float> v1,std::vector<float> v2);

    void CalData(FILE * fileID,std::vector<float> v1,std::vector<float> v2,std::vector<CString> &v_out);
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton10();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton11();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedButton4();


	afx_msg void OnRadioButton_Clicked_1();
	afx_msg void OnRadioButton_Clicked_2();

	afx_msg void OnRadioButton_Clicked_3();
	afx_msg void OnRadioButton_Clicked_4();


	afx_msg void OnComboxButton_Clicked_1();
	afx_msg void OnComboxButton_Clicked_2();
};

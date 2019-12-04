#pragma once


// XHTL 对话框
#include"inputDlg.h"
class XHTL : public CDialogEx
{
	DECLARE_DYNAMIC(XHTL)

public:
	XHTL(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~XHTL();

	CString m_产品代号;
	CString m_boardnum;
	CString m_opname;
	CString m_tempture;
	CString m_humidity;
	CString m_multimeter;
	CString m_constant_volt_source;
	CString m_signal_source;
	CString m_测试工装;
	

	BOOL m_click;

	CApplication app;
    CWorkbooks books;
    CWorkbook book;
    CWorksheets sheets;
    CWorksheet sheet;
    CRange range;
    CRange iCell;
    LPDISPATCH lpDisp;


	CString m_strtime; 


	CString m_EDIT10;
	CString m_EDIT14;
	CString m_EDIT15;
	CString m_EDIT16;
	CString m_EDIT17;
	CString m_EDIT18;
	CString m_EDIT19;
	CString m_EDIT20;
	CString m_EDIT21;
	CString m_EDIT22;


	
    CFont my_Font;


	CBrush m_Brush;

// 对话框数据
	enum { IDD = IDD_XHTL };

	std::vector<float> CStringtoFloat(CString m_s);

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	static DWORD Button1Thread(LPVOID lpParam);
	static DWORD Button2Thread(LPVOID lpParam);
	static DWORD Button3Thread(LPVOID lpParam);
	static DWORD Button4Thread(LPVOID lpParam);



	DECLARE_MESSAGE_MAP()
public:
	afx_msg LRESULT OnUpdateTrueMessage(WPARAM wParam, LPARAM lParam); 
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();
};

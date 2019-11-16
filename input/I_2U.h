#pragma once


// I_2U 对话框
#include "inputDlg.h"
class I_2U : public CDialogEx
{
	DECLARE_DYNAMIC(I_2U)

public:
	I_2U(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~I_2U();

	CString m_boardnum;
	CString m_opname;
	CString m_tempture;
	CString m_humidity;
	CString m_multimeter;
	CString m_constant_volt_source;
	CString m_signal_source;
	CString m_测试工装;
	CString m_测试阶段;

	CString m_strtime;

	CString m_OKorNot;



	CString m_tab2;
	CString m_tab3;
	CString m_tab4;
	CString m_tab5;
	CString m_tab6;
	CString m_tab7;
	CString m_tab8;
	CString m_tab9;
	CString m_tab10;
	CString m_tab11;
	CString m_tab12;

	BOOL m_click;

    CFont my_Font;


	CBrush m_Brush;
	
	
	CApplication app;
    CWorkbooks books;
    CWorkbook book;
    CWorksheets sheets;
    CWorksheet sheet;
    CRange range;
    CRange iCell;
    LPDISPATCH lpDisp;

	std::vector<float> CStringtoFloat(CString m_s);
// 对话框数据
	enum { IDD = IDD_I_2U };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持



	static DWORD Button1Thread(LPVOID lpParam);
	static DWORD Button2Thread(LPVOID lpParam);
	static DWORD Button3Thread(LPVOID lpParam);
	static DWORD Button4Thread(LPVOID lpParam);
	static DWORD Button5Thread(LPVOID lpParam);
    static DWORD Button6Thread(LPVOID lpParam);
	static DWORD Button7Thread(LPVOID lpParam);
    static DWORD Button8Thread(LPVOID lpParam);
	static DWORD Button9Thread(LPVOID lpParam);


	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg LRESULT OnUpdateTrueMessage(WPARAM wParam, LPARAM lParam); 
	afx_msg LRESULT OnUpdateFalseMessage(WPARAM wParam, LPARAM lParam); 
	afx_msg void OnBnClickedButton5();
	afx_msg void OnBnClickedButton6();
	afx_msg void OnBnClickedButton7();
	afx_msg void OnBnClickedButton8();
	afx_msg void OnBnClickedButton9();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedOk();
};

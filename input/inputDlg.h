
// inputDlg.h : 头文件
//


#pragma once
#include<vector>

#include "CApplication.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"


// CinputDlg 对话框
class CinputDlg : public CDialogEx
{
// 构造
public:
	CinputDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_INPUT_DIALOG };

	enum {D_O=0,I2U=1,UI=2,XHTL_ID=3};

	enum {Data_I_2U_ch1=0,Data_I_2U_ch2=1,Data_I_2U_ch3=2,Data_I_2U_ch4=3,Data_I_2U_ch5=4,Data_I_2U_ch6=5};

	CString m_inputstring1;
	CString m_inputstring2;
	CString m_inputstring3;
	CString m_inputstring4;
	CString strFolder;
	int m_filenum;
	FILE * outFileID;
	bool m_getdata;
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

	std::vector<float> CStringtoFloat(CString m_s);

	void getData(char * filename,std::vector<float> v1,std::vector<float> v2);

	void CalData(FILE * fileID,std::vector<float> v1,std::vector<float> v2);
	void CalData_02(FILE * fileID,std::vector<float> v1,std::vector<float> v2,std::vector<float> &out_vec,std::vector<CString> &str_out);


	void EditShowData(CString _s);

	void DrawWave(CDC *pDC,CRect &rectPicture,std::vector<float> x,std::vector<float> y,float a,float b);

	char m_filename[100];
	char m_outfilename[100];
	char m_Rfilename[100];
	char m_historyname[1024];


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

 
    CComboBox m_Select2;
	CComboBox m_Select;
	
     CFont my_Font;


	CBrush m_Brush;

	int m_countclick;//记录输入数据时，对输入进行校验


	BOOL m_saveclick;


	std::vector<float> m_tempv1;
	std::vector<float> m_tempv2;
	float m_tempb1;
	float m_tempb0;

	BOOL m_pictureok;

	void WRfilename(char* filename,bool wr,char *name);
	CString VariantToString(VARIANT var);
// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
	afx_msg void OnBnClickedcolse2();
	afx_msg void OnSelectEvent();
	afx_msg void OnBnClickedcolse3();
};

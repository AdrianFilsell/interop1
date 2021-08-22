
#pragma once

#include "Resource.h"

class TextFieldDlg : public CDialog
{

public:

	// Constructor/Destructor
	TextFieldDlg( CWnd *pParent );
	virtual ~TextFieldDlg();
	
// Dialog Data
	//{{AFX_DATA(TextFieldDlg)
	enum { IDD = IDD_FIELDDIALOG };
	int m_nType;
	CString m_csText;
	CString m_csBtnText;
	CString m_csParaText;
	CString m_csForenameText;
	CString m_csSurnameText;
	//}}AFX_DATA

	// Overrides
	//{{AFX_VIRTUAL(TextFieldDlg)
	virtual void DoDataExchange( CDataExchange* pDX );
	virtual BOOL OnInitDialog( void ) override;
	virtual INT_PTR DoModal( void ) override;
	//}}AFX_VIRTUAL
	
// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(TextFieldDlg)
	afx_msg void OnTypeChange( void );
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	void updatetext( const int nType, const bool bSave );
};

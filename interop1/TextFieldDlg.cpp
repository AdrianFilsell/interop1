
#include "pch.h"
#include "TextFieldDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

TextFieldDlg::TextFieldDlg( CWnd *pParent ):
CDialog( TextFieldDlg::IDD, pParent )
{
	// Init members
	m_nType = 0;

	//{{AFX_DATA_INIT(TextFieldDlg)
	//}}AFX_DATA_INIT
}

TextFieldDlg::~TextFieldDlg()
{
	// Tidy up
}

void TextFieldDlg::DoDataExchange(CDataExchange* pDX)
{
	// Call the base class
	CDialog::DoDataExchange(pDX);
	
	//{{AFX_DATA_MAP(TextFieldDlg)
	DDX_CBIndex( pDX, IDC_FIELDTYPECOMBO, m_nType );
	DDX_Text( pDX, IDC_FIELDTEXTEDIT, m_csText );
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(TextFieldDlg, CDialog)
	//{{AFX_MSG_MAP(TextFieldDlg)
	ON_CBN_SELCHANGE(IDC_FIELDTYPECOMBO,OnTypeChange)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

INT_PTR TextFieldDlg::DoModal( void )
{
	INT_PTR n = CDialog::DoModal();

	updatetext(m_nType,true);

	return n;
}

BOOL TextFieldDlg::OnInitDialog( void )
{
	const BOOL b = CDialog::OnInitDialog();

	updatetext(m_nType,false);

	UpdateData(false);

	return b;
}

void TextFieldDlg::OnTypeChange( void )
{
	const int nType = m_nType;
	UpdateData(true);
	
	updatetext(nType,true);

	updatetext(m_nType,false);

	UpdateData(false);
}

void TextFieldDlg::updatetext( const int nType, const bool bSave )
{
	if( bSave )
		switch( nType )
		{
			case 0:m_csBtnText = m_csText;break;
			case 1:m_csParaText = m_csText;break;
			case 2:m_csForenameText = m_csText;break;
			case 3:m_csSurnameText = m_csText;break;
			default:ASSERT(false);return;
		}
	else
		switch( nType )
		{
			case 0:m_csText = m_csBtnText;break;
			case 1:m_csText = m_csParaText;break;
			case 2:m_csText = m_csForenameText;break;
			case 3:m_csText = m_csSurnameText;break;
			default:ASSERT(false);return;
		}
}

/*------------------------------------------------------------------------------*
 * File Name:				 													*
 * Creation: 																	*
 * Purpose: OriginC Source C file												*
 * Copyright (c) ABCD Corp.	2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010		*
 * All Rights Reserved															*
 * 																				*
 * Modification Log:															*
 *------------------------------------------------------------------------------*/
 
////////////////////////////////////////////////////////////////////////////////////
// Including the system header file Origin.h should be sufficient for most Origin
// applications and is recommended. Origin.h includes many of the most common system
// header files and is automatically pre-compiled when Origin runs the first time.
// Programs including Origin.h subsequently compile much more quickly as long as
// the size and number of other included header files is minimized. All NAG header
// files are now included in Origin.h and no longer need be separately included.
//
// Right-click on the line below and select 'Open "Origin.h"' to open the Origin.h
// system header file.
#include <Origin.h>
////////////////////////////////////////////////////////////////////////////////////

//#pragma labtalk(0) // to disable OC functions for LT calling.

////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.
#include <..\OriginLab\DialogEx.h>
//#include "OdlgA.h"

#include <ocGDI.h>
#include "TemplateLibraryDlg.h"

#include <ocu.h>	
#include <XFbase.h>	
#include "grobj_utils.h"


#define STR_USER_PATH 			okutil_get_origin_path(ORIGIN_PATH_USER, NULL, TRUE)

#define NUM_CATEGORY_ICON_PIXEL					16
#define NUM_CATEGORY_LIST_WIDTH					170

#define NUM_TEMPLATE_ICON_PIXEL					16//48
#define NUM_TEMPLATE_ICON_WIDTH					850
#define NUM_TEMPLATE_ICON_LIST_COLS				6

#define NUM_TEMPLATE_PREVIEW_WIDTH				2000
#define NUM_TEMPLATE_PREVIEW_LIST_COLS			4

#define NUM_CATEGORY_LABEL_CELL_COLOR			RGB(0xF0,0xF0,0xF0)//RGB(146, 188, 211)


#define	SEL_CELL_BG_COLOR						RGB(200, 225, 250)
#define	SEL_CELL_FG_COLOR						RGB(0, 0, 0)

/// Iris 4/13/2015 ORG-12495-S8 IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE
#define	RGB_WHITE_COLOR							RGB(255, 255, 255)
///End IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE


#define	STR_MOUSEOVERGRAY_EMF_FILENAME	(GetAppPath(TRUE) + "Themes\\TLDToolbarGray.emf")
#define	STR_MOUSEOVER_EMF_FILENAME		(GetAppPath(TRUE) + "Themes\\TLDToolbarColor.emf")


#define	STR_INDCATOR_EMF_FILENAME		(GetAppPath(TRUE) + "Themes\\Dolly.emf")

#define	STR_INDCATOR_EMF_FILENAME_GRAY		(GetAppPath(TRUE) + "Themes\\DollyGray.emf")	///------ Tony 08/18/2015 ORG-13290-S9 GRAY_SHEEP_IF_NOT_MATCH

#define STR_PICT_NO_PREVIEW						"NO_PREVIEW"
#define STR_PICT_NON_TEMPLATE_HINT				"nonTemplateHint"
#define STR_PICT_NO_TEMPLATE_MATCH_HINT			"NoMatchHint" /// Iris 6/3/2015 ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER

#define STR_ORIGIN_USER_FILE_FOLDER				GetOriginPath(ORIGIN_PATH_USER)
#define STR_MODULE_DLL_FILE_NAME				ORIGINC_SOURCE_PATH_ORIGINLAB + "ODlgA"

#define STR_CATEGORYS							"User Defined|Line|Symbol|Line + Symbol|Column/Bar/Pie|Multi-Y|Y-offset/Waterfall|Multi-Panel|3D XYY|3D Surface|3D Symbol/Bar/Vector|Statistics|Area|Contour/Heat Map|Profile|Specialized|Stock"
#define STR_NEW_CATEGORY						"New Category"

typedef BOOL (*FUNC_PLOT_SETUP)(BOOL bNewPlot, int nPlotID, LPCTSTR lpcszTemplate, HWND hWndParent );


#define MOUSE_OVER_ICON_NUM 2

enum
{
	TOOL_BAR_INDEX_EDIT = 0,
	TOOL_BAR_INDEX_DELETE
};

////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

//FilterGraphTemplate is slow, so remember it
static string s_strTemplate = "";
static string s_strSheetActive = "";
static DWORD s_dwTemplate = 0;
static void _clear_filter_graph_template()
{
	s_strTemplate = "";
	s_strSheetActive = "";
	s_dwTemplate = 0;
}

static DWORD _filter_graph_template(string strTemplate, string strSheetActive)
{
	if(strTemplate.IsEmpty())
		return s_dwTemplate;
	
	bool bNeedGet = false;
	
	if(s_strTemplate.IsEmpty() || 0 != s_strTemplate.CompareNoCase(strTemplate) )
	{
		bNeedGet = true;
		s_strTemplate = strTemplate;
	}
	
	if(bNeedGet)
	{
		s_strSheetActive = strSheetActive;
	}
	else if( !strSheetActive.IsEmpty() && 0 != s_strSheetActive.CompareNoCase(strSheetActive) )
	{
		bNeedGet = true;
		s_strSheetActive = strSheetActive;
	}
	
	if(bNeedGet)
	{
		LPCSTR lpSheetActive = NULL;
		if(!s_strSheetActive.IsEmpty())
			lpSheetActive = s_strSheetActive;
	
		s_dwTemplate = Project.FilterGraphTemplate(s_strTemplate, lpSheetActive);	
	}
	
	return s_dwTemplate;
}

static bool _is_template_cloneable(LPCSTR lpcszTemplate)
{
	if ( NULL == lpcszTemplate || '\0' == *lpcszTemplate )
		return false;
	
	DWORD dw = _filter_graph_template(lpcszTemplate, "");
	
	return !O_QUERY_BOOL(dw, FilterTemplateRst_NotCloneableTemplate);		
}

static bool _is_template_match_with_active_sheet(LPCSTR lpcszTemplate)
{
	Page pg = Project.Pages();
	
	if( !pg || !_is_page_support_auto_plotting(pg.GetType()) )
		return false;
	
	Datasheet sheetActive;
	sheetActive = pg.Layers();
	DWORD dw = _filter_graph_template(lpcszTemplate, wks_get_book_sheet_name(sheetActive));
	
	return is_template_match_ok(dw);
}

static TemplateLibraryDlg* s_pTemplateLibraryDlg;
static string	s_strUpdatedTemplateFile;


static bool s_bOnActivateReady = true;

static int _MessageBox(HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType)
{
	bool bSet = false;
	if(s_bOnActivateReady)
	{
		s_bOnActivateReady = false;
		bSet = true;
	}
	
	int nRet = MessageBox(hWnd, lpText, lpCaption, uType);
	
	if(bSet)
		s_bOnActivateReady = true;
	
	return nRet;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////  TemplateListBase  /////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
void TemplateListBase::AddTextRow(int nRow, LPCSTR lpcstrText, bool bHideMouseOverIcon /*= false*/)
{
	if( nRow == GetNumRows() )
		AddOneRow();
	for(int nCol = 0; nCol < GetNumCols(); nCol++)
	{
		SetCell(nRow, nCol, lpcstrText);
		SetCellProperty(flexcpBackColor, NUM_CATEGORY_LABEL_CELL_COLOR, nRow, nCol);
		SetBold(nRow, nCol);
		SetCellHideMouseOverIcon(nRow, nCol, bHideMouseOverIcon);
	}
	SetMergeRow(nRow, true);
	SetRowHeight(nRow, GetRowHeight(nRow) * 1.2 );
}

virtual	void TemplateListBase::Init(int nID, WndContainer& dlg, TemplateSettingsHelper* pXMLHelper)
{
	ASSERT(NULL != pXMLHelper);
	m_pXMLHelper = pXMLHelper;
	
	GridListControl::Init(nID, dlg);
	
	SetExtendLastCol(false);
	SetupRowsCols(0);
	SetSelection(flexSelectionFree);
	SetDragImage(true);
	SetDragCell(true);
	
    SetCellHighlightEnable(true);
	
}

void TemplateListBase::SetCellHighlightEnable(bool bEnable)
{
	m_flx.HighlightFocusedCell = bEnable;
	if( bEnable )
	{
		m_flx.BackColorSel = SEL_CELL_BG_COLOR;
		m_flx.ForeColorSel = SEL_CELL_FG_COLOR;
	}
	else
	{
		m_flx.BackColorSel = RGB_WHITE_COLOR;
		m_flx.ForeColorSel = RGB_WHITE_COLOR;
	}
}

bool TemplateListBase::IsCategoryRow(int nRow)
{
	return IsMergeRow(nRow);
}

void TemplateListBase::SetupScrollBarPos(HWND hWnd, int nRow)
{
	if( !IsReady() || nRow < 0 || NULL == hWnd )
		return;		
	
	Window wndGrid(hWnd);
	wndGrid.SetScrollPos(SB_VERT, 0, 100);
	
	int min, max;	
	wndGrid.GetScrollRange(SB_VERT, min, max);
	//printf("min=%d, max=%d\r\n", min, max);
	
	int nItemCount = (CATEGORY_ITEM_COUNT - CATEGORY_LINE - 1);
	int nOneSetup = (max - min) / nItemCount;
	int nPos = (nRow-1) * nOneSetup;
	if( nRow > 1 && nRow + 1 < nItemCount )
		nPos -= nOneSetup;
	_DBINT("nPos=", nPos);
	
	if( nPos < min )
		nPos = min;
	if( nPos > max )
		nPos = max;
	
	SetScrollBarPos(hWnd, nPos, TRUE);
}

void TemplateListBase::SetScrollBarPos(HWND hWnd, int nPos, bool bRefresh)
{
	Window wndGrid(hWnd);
	if( wndGrid )
	{	
		wndGrid.SetScrollPos(SB_VERT, nPos, bRefresh);	
		
		DWORD wParam = MAKEWPARAM(SB_THUMBTRACK, nPos);	
		wndGrid.PostMessage(WM_VSCROLL, wParam, 0);
	}
}

double TemplateListBase::GetScrollBarPosPercent(HWND hWnd)
{
	Window wndGrid(hWnd);
	if( wndGrid )
	{
		double dPos = wndGrid.GetScrollPos(SB_VERT);
		
		int min, max;
		wndGrid.GetScrollRange(SB_VERT, min, max);
		
		_DBPRINTF("pos = %g\n", dPos);
		_DBPRINTF("get pos (%) = %g\n", (dPos / (max - min) * 100));
		return (dPos / (max - min) * 100);
	}
	
	return 0;
}

void TemplateListBase::SetScrollBarPosPercent(HWND hWnd, double dPosPercent, bool bRefresh)
{
	Window wndGrid(hWnd);
	if( wndGrid )
	{	
		int min, max;	
		wndGrid.GetScrollRange(SB_VERT, min, max);
		
		int nPos = (int)(max - min) * dPosPercent / 100;
		
		SetScrollBarPos(hWnd, nPos, bRefresh);
	}
}

bool TemplateListBase::BeforeLoad(HWND hWndList)
{		
	if( GetNumRows() > 0 && NULL != hWndList )	
	{		
		// need to restore the position of scroll bar
		m_dScrollBarPosPercent = GetScrollBarPosPercent(hWndList);
		
		return true; // return true means reload the list
	}
	return false; 
}

void TemplateListBase::AfterLoad(HWND hWndList, bool bReload)
{	
	if( bReload )	
	{
		_DBPRINTF("scroll bar pos = %g\n", m_dScrollBarPosPercent);
		if( 0 != m_dScrollBarPosPercent )
			SetScrollBarPosPercent(hWndList, m_dScrollBarPosPercent);
	}
}

// virtual
int TemplateListBase::GetCategoryIndex(int nRow)
{	
	if( nRow < 0 )	
	{
		ASSERT(FALSE);
		return 0;
	}
	
	int count = 0;
	while( nRow >= GetRowOffset() )
	{
		if( IsCategoryRow(nRow) )
		{			
			count++;
		}
		nRow--;
	}
		
	if( m_pXMLHelper->HasUntitledCategory() && !IsCategoryRow(0) )
		count++;
		
	ASSERT( count > 0 );

	return count - 1;
}

void	TemplateListBase::SetupMouseOverIcon(int nHighlight)
{
	string strEMFFile = 0 == nHighlight ? STR_MOUSEOVERGRAY_EMF_FILENAME : STR_MOUSEOVER_EMF_FILENAME;
	HENHMETAFILE hEnEMF = load_enhanced_metafile(strEMFFile);	
	PictureHolder pict;
	if ( pict.CreateFromEnhMetafile(hEnEMF, false) )
	{
		SetMouseOverIcon(pict, nHighlight);
		SetMouseOverIconNum(MOUSE_OVER_ICON_NUM);
	}
}

void	TemplateListBase::SetupCellIndicator(int nHighlight)
{
	string strEMFFile = nHighlight ? STR_INDCATOR_EMF_FILENAME : STR_INDCATOR_EMF_FILENAME_GRAY;
	
	HENHMETAFILE hEnEMF = load_enhanced_metafile(strEMFFile);	
	PictureHolder pict;
	if ( pict.CreateFromEnhMetafile(hEnEMF, false) )
    {
		
        SetCellIndicatorIcon(pict, nHighlight);
		
	}
}


void TemplateListBase::OnDlgResize(int cx, int cy)
{
	if( GetNumCols() != CalcColNum() )
		Load();
}

int TemplateListBase::CalcColNum()
{
	RECT rrList;
	GetWindowRect(rrList);
	
	int iTwips = PixelsToTwips( RECT_WIDTH(rrList) );
	double dNumCols = iTwips / getCellWidth();
	int nNumCols = (int)dNumCols;
	
	if( nNumCols < 4 )
		nNumCols = 4;
	return nNumCols;
}

static bool _load_image_to_picture_holder(LPCSTR lpcszModuleFile, int nBitmapID, PictureHolder& pict)
{
	HMODULE hModule = okutil_load_dll(lpcszModuleFile);
	if( NULL == hModule )
	{
		error_report("Fail to load module dll.");
		return false;
	}	
	
	HBITMAP hBmp = LoadBitmap(hModule, (LPCSTR)nBitmapID);	
	if( !pict.CreateFromBitmap(hBmp, TRUE) )
	{
		error_report("Fail to load template category list image from module.");
		return false;
	}	
	
	return true;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////  PreviewList  ///////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
static HWND s_hWndPreviewList = NULL;

static BOOL GetPreviewWndCallback(HWND hWnd, LPARAM lParam)
{
	s_hWndPreviewList = hWnd;
	return FALSE;
}

void PreviewList::Init(int nID, WndContainer& dlg, TemplateSettingsHelper* pXMLHelper, LPCSTR lpcszDefaultTemplate)
{
	TemplateListBase::Init(nID, dlg, pXMLHelper);
	
	Control* pctrl = GetControl();
	if( pctrl )
	{		
		s_hWndPreviewList = NULL;
		EnumChildWindows(pctrl->GetSafeHwnd(), GetPreviewWndCallback, 0);		
	}
	
	m_strDefaultTemplate = lpcszDefaultTemplate;
	SetReady();
}

bool PreviewList::loadImage(LPCSTR lpcszFilename, int nRow, int nCol, bool bShowTemplateName)
{
	string strFilename(lpcszFilename);
	
	if( strFilename.IsEmpty() || !strFilename.IsFile() )
		return false;
	
	bool bIsBMP, bIsEMF;
	string strImageFile = GetFilePath(strFilename) + GetFileName(strFilename, true) + ".bmp";
	if( strImageFile.IsFile() )
		bIsBMP = true;
	else
	{
		strImageFile = GetFilePath(strFilename) + GetFileName(strFilename, true) + ".emf";
		bIsEMF = strImageFile.IsFile();
	}	
	
	if( !strImageFile.IsFile() )
	{
		GetTempFileName(strImageFile);
		strImageFile += "preview.png";
		int nExtractRet = okutil_extract_Origin_file_preview(strFilename, strImageFile, 1, true);
		if (nExtractRet > 0)
			bIsBMP = TRUE;
	}
	
	/// Iris 4/13/2015 ORG-12495-S8 IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE
	/*
	if( !strImageFile.IsFile() )
		return false;
	*/
	if( !strImageFile.IsFile() )
	{
		strImageFile = GetFullPath(STR_PICT_NO_PREVIEW ".emf", "Templates\\Previews", TRUE);
		ASSERT( strImageFile.IsFile() );
		bIsEMF = strImageFile.IsFile();
	}
	///End IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE		
	
	bool bRet = false;
	PictureHolder pict;
	if( bIsEMF )
	{
		HENHMETAFILE hEnEMF = load_enhanced_metafile(strImageFile);			
		bRet = pict.CreateFromEnhMetafile(hEnEMF, false);			
	}
	else if( bIsBMP )
	{			
		bRet = pict.Load(strImageFile);
	}
	else
	{
		// should load No Preview image, to do 
	}	
	if( !bRet )
		return error_report("fail to load image file - " + strImageFile);
	
	/// Iris 5/8/2015 ORG-12996-P3 FIX_INCORRECT_IMAGE_TITLE
	//setImage(nRow, nCol, pict, strImageFile);
	setImage(nRow, nCol, pict, strFilename, bShowTemplateName);		
	///End FIX_INCORRECT_IMAGE_TITLE
	return true;
}

bool	PreviewList::OnTemplateUpdate(string strTemplate)
{
	if ( m_pXMLHelper && m_pXMLHelper->CheckCategory(strTemplate) ) //already exist, need only update the cell
	{
		int nRows = GetRows();
		int nCols = GetCols();
		for ( int iRow = GetRowOffset(); iRow < nRows; iRow++ )
		{
			for ( int iCol = GetColOffset(); iCol < nCols; iCol++ )
			{
				string strFileName = GetCell(flexcpData, iRow, iCol);
				if ( strFileName.IsEmpty() )
					break; //not any more template in this category, check next category
				if ( strTemplate.CompareNoCase(strFileName) == 0 )
				{
					return loadImage(strTemplate, iRow, iCol);
				}
				
			}
		}
	}
	else //better reconstruct all on newly saved template noticed.
	{
		Load(true);
	}
	return true;
}
void	PreviewList::OnActivate()
{
	int nRows = GetRows();
	int nCols = GetCols();
	for ( int iRow = GetRowOffset(); iRow < nRows; iRow++ )
	{
		for ( int iCol = GetColOffset(); iCol < nCols; iCol++ )
		{
			string strFileName = GetCell(flexcpData, iRow, iCol);
			if ( strFileName.IsEmpty() )
				break; //next row
			if ( strFileName.IsFile() )
			{
				bool bMatch = _is_template_match_with_active_sheet(strFileName);
				bool bCloneable = _is_template_cloneable(strFileName);
				int nPosition = bCloneable ? (bMatch ? flexIndicatorPositionLeftBottomHighlight : flexIndicatorPositionLeftBottom) : flexIndicatorPositionNone;
				SetCell(nPosition, flexcpPictureIndicatorPosition, iRow, iCol);
			}
		}
	}
}

void PreviewList::SetIsTemplateNone(bool bIsTemplateNone)
{
	m_bIsTemplateNone = bIsTemplateNone;
}

bool PreviewList::IsTemplateNone()
{
	return m_bIsTemplateNone;
}

bool PreviewList::SetFilterOption(bool bIsFilterOn)
{
	bool bUpdate = false;
	
	/// Iris 5/27/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTONS
	if( m_bIsFilterOn != bIsFilterOn )
	{
		m_bIsFilterOn = bIsFilterOn;
		bUpdate = true;
	}
	///End IMPROVE_PLOT_BUTTONS
	
	return bUpdate;
}


bool PreviewList::IsFilterOn()
{		
	return m_bIsFilterOn;
}

bool PreviewList::filter(LPCSTR lpcszTemplate)
{
	Page pg = Project.Pages();

	if( !pg || !_is_page_support_auto_plotting(pg.GetType()) )
		return true; // or should return false??	
	
	Datasheet sheetActive;
	sheetActive = pg.Layers();
	/// Iris 6/25/2015 ORG-12495-S10 IMPROVE_TEMPLATE_FILTER
	//DWORD dw = Project.FilterTemplate(lpcszTemplate, wks_get_book_sheet_name(sheetActive), dwMatchCtrl, IDM_PLOT_UNKNOWN);
	DWORD dw = Project.FilterGraphTemplate(lpcszTemplate, wks_get_book_sheet_name(sheetActive));
	///End IMPROVE_TEMPLATE_FILTER
	
	/// Iris 4/23/2015 ORG-12495-S1 FIX_CROSS_SHEET_TEMPLATE_FIAL_TO_FILTER
	/// Iris 6/1/2015 ORG-12495-P1 FIX_DATA_FROM_AUTO_FAILT_TO_DO_FILTER
	/// Iris 6/2/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTONS
	//if( O_QUERY_BOOL(dw, FilterTemplateRst_SrcCrossSheet) ? FILTER_BY_WIN_SHEET == m_nDataFrom : FILTER_BY_WIN_SHEET != m_nDataFrom )
	//	return false;
	/// Iris 6/25/2015 ORG-12495-S10 IMPROVE_TEMPLATE_FILTER
	/*
	if( FILTER_BY_WIN_AUTO != m_nDataFrom )
	{
		if( O_QUERY_BOOL(dw, FilterTemplateRst_SrcCrossSheet) ? FILTER_BY_WIN_BOOK != m_nDataFrom : FILTER_BY_WIN_BOOK == m_nDataFrom )
			return false;
	}
	*/
	///End IMPROVE_TEMPLATE_FILTER
	///End IMPROVE_PLOT_BUTTONS
	///End FIX_DATA_FROM_AUTO_FAILT_TO_DO_FILTER
	///End FIX_CROSS_SHEET_TEMPLATE_FIAL_TO_FILTER
	
	if( !O_QUERY_BOOL(dw, FilterTemplateRst_HasSrcMatched)  // not auto plotting info match in StyleHolder		
	 || dw >= FilterTemplateRst_Error 
		|| O_QUERY_BOOL(dw, FilterTemplateRst_HasAxisSeriesSkipped)		///Kyle 04/29/2015 ORG-12522-P3 AUTO_PLOT_FROM_TEMPLATE_NEED_CHANGE_SERIES_USED_IN_AXES
		|| O_QUERY_BOOL(dw, FilterTemplateRst_HasSrcSkipped) 
		/// Iris 4/23/2015 ORG-12495-S1 FIX_CROSS_SHEET_TEMPLATE_FIAL_TO_FILTER
		//|| O_QUERY_BOOL(dw, FilterTemplateRst_SrcCrossSheet)
		///End FIX_CROSS_SHEET_TEMPLATE_FIAL_TO_FILTER
	)
	{	
		return false;
	}	
	return true;
}

/// Iris 6/25/2015 ORG-12495-S10 IMPROVE_TEMPLATE_FILTER
//bool PreviewList::filter(LPCSTR lpcszTemplate)
//{
	//Page pg = Project.Pages();
	//if( !pg || !( EXIST_WKS == pg.GetType() || EXIST_MATRIX == pg.GetType() ) )
		//return true; // or should return false??	
	//
	//Datasheet sheetActive;
	//sheetActive = pg.Layers();
	//DWORD dw = Project.FilterTemplate(lpcszTemplate, wks_get_book_sheet_name(sheetActive));	
	//
	//if( !O_QUERY_BOOL(dw, FilterTemplateRst_HasSrcMatched)  // not auto plotting info match in StyleHolder		
	 //|| dw >= FilterTemplateRst_Error 
		//|| O_QUERY_BOOL(dw, FilterTemplateRst_HasAxisSeriesSkipped)		///Kyle 04/29/2015 ORG-12522-P3 AUTO_PLOT_FROM_TEMPLATE_NEED_CHANGE_SERIES_USED_IN_AXES
		//|| O_QUERY_BOOL(dw, FilterTemplateRst_HasSrcSkipped)		
	//)
	//{
		//return false;
	//}	
	//return true;
//}
///End IMPROVE_TEMPLATE_FILTER

bool PreviewList::CheckByFilter(LPCSTR lpcszTemplate, int nCategoryIndex/* = -1*/, int nTemplateIndex/* = -1*/)
{
	/// Iris 6/25/2015 ORG-12495-S10 IMPROVE_TEMPLATE_FILTER
	//GraphPage gp;
	//if( !gp.Create( lpcszTemplate, CREATE_HIDDEN | CREATE_SET_MISSING_IN_MANAGER) || !gp )
		//return error_report("fail to do filter since fail to create graph window with template.");
	//
	//bool bMatch = gp.IsCloneableEnabled() && filter( lpcszTemplate );	
	//gp.Destroy();
	//
	//if( !bMatch )
		//return false; // not included clone info or not match, no need to following.
	
	if( !_is_template_match_with_active_sheet(lpcszTemplate) )
		return false;
	///End IMPROVE_TEMPLATE_FILTER
	
	if( nCategoryIndex >= 0 && nTemplateIndex >= 0 )
	{
		TreeNode trCategory = tree_check_get_node(m_trFilter, "Category" + (nCategoryIndex + 1));
		///------ Tony 07/14/2015 ORG-12495-S10 IMPROVE_TEMPLATE_FILTER
		TreeNode tr = tree_get_node_by_id(trCategory, nTemplateIndex);
		if(!tr)
		///------ End IMPROVE_TEMPLATE_FILTER
		{
			trCategory.DataID = nCategoryIndex;
		
			string strCategoryLabel;
			m_pXMLHelper->GetCategoryLabel(nCategoryIndex, strCategoryLabel);
			trCategory.SetAttribute(STR_LABEL_ATTRIB, strCategoryLabel);

			TreeNodeCollection trColl(trCategory, "Template");
			string strTagname = "Template" + ( trColl.Count() + 1 );
			TreeNode trTemplate = trCategory.AddNode(strTagname);
			trTemplate.DataID = nTemplateIndex;
		
		}
	}
	
	return true;
}

///------ Tony 09/29/2015 ORG-12495-P9 SAVE_XML_TREE_WHEN_INACTIVATE
//bool PreviewList::Load()
bool PreviewList::Load(bool bReloadXML /*false*/)
///------ End SAVE_XML_TREE_WHEN_INACTIVATE
{	
	_DBSTR("PreviewList::Load()");	

	//template has be changed, so should clear
	_clear_filter_graph_template();

	///------ Tony 09/29/2015 ORG-12495-P9 SAVE_XML_TREE_WHEN_INACTIVATE
	if(bReloadXML)
		m_pXMLHelper->LoadXML();
	///------ End SAVE_XML_TREE_WHEN_INACTIVATE
	
	bool bReload = BeforeLoad(s_hWndPreviewList);
	
	// if reload all template preview images, need to restore the cell selection.
	string strTemplateToSel;
	if( bReload )
	{
		int nRow, nCol;
		if( GetSelCell(&nRow, &nCol) && nRow >= 0 && nCol >= 0 )			
			strTemplateToSel = GetTemplateFilePath(nRow, nCol);		
	}
	/// Iris 5/27/2015 ORG-12495-S11 OPEN_TEMPLATE_LIB_DLG_FROM_PLOT_MENU
	else if( !m_strDefaultTemplate.IsEmpty() )
	{
		strTemplateToSel = m_strDefaultTemplate;
		m_strDefaultTemplate.Empty(); // only use it once time
	}
	///End OPEN_TEMPLATE_LIB_DLG_FROM_PLOT_MENU
	
	// Reset 
	///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	//SetRows(0);
	//SetCols(CalcColNum());	
	//set row or col will activate the selChange() event, and load() will be called when GetNumCols() != CalcColNum(), so should set the cols first, then load() will be called once
	SetCols(CalcColNum());	
	SetRows(0);
	///------ End IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	SetColAlignment(-1, flexAlignLeftCenter);	
	m_trFilter.Reset();
	
	bool bSelCellDone = false;
	int nRowToSel = -1, nColToSel = -1;
	
	// loading...
	SetDrawPictureAsBackground(true);
	if(!CheckLoadImageForTempalteNone())
	{
	    int nRow = 0, nCol = 0;
	
		for(int nCategoryIndex = 0; nCategoryIndex < m_pXMLHelper->GetCategoryCount(); nCategoryIndex++)
		{	
			int nTemplateCount = m_pXMLHelper->GetTemplateCount(nCategoryIndex);
			
			// add the category name to one row.
			string strCategoryLabel;		
			m_pXMLHelper->GetCategoryLabel(nCategoryIndex, strCategoryLabel);			
			/// Iris 4/22/2015 ORG-12495-S7 SHOW_USER_DEFINED_CATEGORY_LABEL
			/// Iris 5/29/2015 ORG-12495-S7 USER_DEFINED_CATEGORY_SHOULD_BE_NORMAL_CATEGORY
			//bool bIsTempateWithoutCategorlyLabel = ( 0 == strCategoryLabel.CompareNoCase(STR_USER_DEFINED_CATEGORY_LABEL) );
			bool bIsTempateWithoutCategorlyLabel = ( 0 == strCategoryLabel.CompareNoCase(STR_UNTITLED_CATEGORY_LABEL) );
			///End USER_DEFINED_CATEGORY_SHOULD_BE_NORMAL_CATEGORY
			///End SHOW_USER_DEFINED_CATEGORY_LABEL
			
			/// Iris 4/22/2015 ORG-12495-S7 SHOW_USER_DEFINED_CATEGORY_LABEL
			//if( !strCategoryLabel.IsEmpty() && !bIsTempateWithoutCategorlyLabel )
			/// Iris 5/29/2015 ORG-12495-S7 USER_DEFINED_CATEGORY_SHOULD_BE_NORMAL_CATEGORY
			//if( !strCategoryLabel.IsEmpty() )
			if( !strCategoryLabel.IsEmpty() && !bIsTempateWithoutCategorlyLabel )
			///End USER_DEFINED_CATEGORY_SHOULD_BE_NORMAL_CATEGORY
			///End SHOW_USER_DEFINED_CATEGORY_LABEL
			{		
				AddTextRow(nRow, strCategoryLabel);
				
				// go to next row to put preview images
				nRow++;
				nCol = 0;
			}
			
			int nNumImagesShow = 0;
			if( nTemplateCount > 0 )
			{
				int nRowBegin = nRow;
				int nImageIndex = 0;
				
				for(int nTemplateIndex = 0; nTemplateIndex < nTemplateCount; nTemplateIndex++)
				{		
					string strTemplateFile;
					m_pXMLHelper->GetTemplateFilePath(nCategoryIndex, nTemplateIndex, strTemplateFile);					
					
					if( IsFilterOn() && !CheckByFilter(strTemplateFile, nCategoryIndex, nTemplateIndex) )					
						continue;					
					
					nRow = nRowBegin + nImageIndex / GetNumCols();
					nCol = mod(nImageIndex, GetNumCols());
					
					// load image into grid cell
					checkAddImageRow(nRow);
					/// Iris 6/11/2015 ORG-13203-P2 FIX_FAIL_TO_DISPLAY_NOPREVIEW_HINT_EMF_WHEN_TEMPALTE_NO_IMAGE_FILE
					//loadImage(strImageFile, nRow, nCol);
					loadImage(strTemplateFile, nRow, nCol);
					///End FIX_FAIL_TO_DISPLAY_NOPREVIEW_HINT_EMF_WHEN_TEMPALTE_NO_IMAGE_FILE
					
					// set row selection
					if( !strTemplateToSel.IsEmpty() && 0 == strTemplateFile.CompareNoCase(strTemplateToSel) )
					{
						SelCell(nRow, nCol);
						bSelCellDone = true;
						nRowToSel = nRow;
						nColToSel = nCol;
					}

					SetRowHeight(nRow, getCellWidth());
					SetColWidth(nCol, getCellWidth());
					nImageIndex++;
					nNumImagesShow++;
				}	
				
				if( nImageIndex > 0 )
					nRow++; // go to next category
			}
			else
			{
				// if the current category is not Untitled category, then add an empty image row for it.				
				if( !bIsTempateWithoutCategorlyLabel )	
				{
					checkAddImageRow(nRow);
					nRow++; // go to next category
				}				
			}			
			///------ Tony 07/23/2015 ORG-13290-S5 INITIALIZE_THE_CATEGORY_NAME_IMPROVE
			//if( IsFilterOn() && !bIsTempateWithoutCategorlyLabel && 0 == nNumImagesShow ) // not dispaly empty category when filter is ON			
			if( IsFilterOn() && !bIsTempateWithoutCategorlyLabel && 0 == nNumImagesShow && !IsCategoryRow(nRow - 1)) // not dispaly empty category but keep the categroy row when filter is ON			
			///------ End INITIALIZE_THE_CATEGORY_NAME_IMPROVE
			{
				nRow--;		
				if( nRow >= 0 )
					DeleteRow(nRow); // to delete the content row
			}
		}
		/// Iris 6/3/2015 ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
		///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
		//if( isFilterOn() && 0 == GetRows() )
		//{
			//LoadImageForTempalteNone(STR_PICT_NO_TEMPLATE_MATCH_HINT);
			//m_bIsTemplateNone = true;
		//}
		///End ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
		///------ End IMPROVE_HINT_IMAGE
	}
	
	/// Iris 4/13/2015 ORG-12495-S8 IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE
	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	//if( !m_bIsTemplateNone )
	if( !IsTemplateNone() )
	///------ End IMPROVE_HINT_IMAGE
	{
		///Sophy 3/3/2015 ORG-12495-S3 SUPPORT_OVERLAY_ICON_ON_CELL_WHEN_MOUSE_OVER for test, load any
		SetupMouseOverIcon(0);
		SetupMouseOverIcon(1);
		///end SUPPORT_OVERLAY_ICON_ON_CELL_WHEN_MOUSE_OVER	
		///------ Tony 08/18/2015 ORG-13290-S9 GRAY_SHEEP_IF_NOT_MATCH
		//SetupCellIndicator();	///Sophy 7/9/2015 ORG-12495-S10 IMAGE_AS_INDICATOR_AT_CELL_BOTTOM_LEFT
		SetupCellIndicator(0);
		SetupCellIndicator(1);
		///------End GRAY_SHEEP_IF_NOT_MATCH

	}
	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	//SetCellHighlightEnable( !m_bIsTemplateNone );
	SetCellHighlightEnable( !IsTemplateNone() );
	///------ End IMPROVE_HINT_IMAGE
	///End IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE
	
	if( !IsTemplateNone() )	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	{
		// if no cell be selected yet, default should select the first template.
		SetReady(false);
		if( !bSelCellDone )
		{
			/// Iris 6/11/2015 ORG-13203-P2 FIX_FAIL_TO_DISPLAY_NOPREVIEW_HINT_EMF_WHEN_TEMPALTE_NO_IMAGE_FILE
			//SelCell(1, 0); // the first template in the first category.
			int nRow = IsCategoryRow(0) ? 1 : 0;
			SelCell(nRow, 0); // the first template in the first category.
			///End FIX_FAIL_TO_DISPLAY_NOPREVIEW_HINT_EMF_WHEN_TEMPALTE_NO_IMAGE_FILE
		}
		else
			SelCell(nRowToSel, nColToSel);	
		SetReady(true);
	}
	
	AfterLoad(s_hWndPreviewList, bReload);	
	return true;
}

///------ Tony 09/29/2015 ORG-12495-P9 SAVE_XML_TREE_WHEN_INACTIVATE
bool PreviewList::Save()
{
	m_pXMLHelper->Save();
	
	return true;
}
///------ End SAVE_XML_TREE_WHEN_INACTIVATE

///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
bool PreviewList::CheckLoadImageForTempalteNone()
{
	m_bIsTemplateNone = false;
	if( IsFilterOn() )
	{	
		bool bHasTemplate = false;
		for(int nCategoryIndex = 0; nCategoryIndex < m_pXMLHelper->GetCategoryCount(); nCategoryIndex++)
		{
			int nTemplateCount = m_pXMLHelper->GetTemplateCount(nCategoryIndex);
			
			for(int nTemplateIndex = 0; nTemplateIndex < nTemplateCount; nTemplateIndex++)
			{		
				string strTemplateFile;
				m_pXMLHelper->GetTemplateFilePath(nCategoryIndex, nTemplateIndex, strTemplateFile);					
				
				if( IsFilterOn() && !CheckByFilter(strTemplateFile, nCategoryIndex, nTemplateIndex) )									
					continue;	
				
				bHasTemplate = true;
				break;
			}
		}
		
		if( !bHasTemplate )
		{
			LoadImageForTempalteNone(STR_PICT_NO_TEMPLATE_MATCH_HINT);
			
			return true;
		}
	}
	else if( 0 == m_pXMLHelper->GetCategoryCount() )
	{
		LoadImageForTempalteNone(STR_PICT_NON_TEMPLATE_HINT);

		return true;
	}
	
	
	return false;
}
///------ End IMPROVE_HINT_IMAGE

void PreviewList::AddTextRow(int nRow, LPCSTR lpcstrText)
{
	TemplateListBase::AddTextRow(nRow, lpcstrText, IsFilterOn());
}

/// Iris 4/13/2015 ORG-12495-S8 IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE
/// Iris 6/3/2015 ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
//bool PreviewList::LoadImageForTempalteNone()
bool PreviewList::LoadImageForTempalteNone(LPCSTR lpcszImageFileName)
///End ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
{
	m_flx.HighlightFocusedCell = false;
	m_flx.BackColorSel = RGB(255, 255, 255);
	m_flx.ForeColorSel = RGB(255, 255, 255);
	
	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	SetIsTemplateNone(true);
	//clean up first
	SetRows(0);
	///------ End IMPROVE_HINT_IMAGE
	
	SetRows(1);
	SetCols(1);
	
	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	SetCellHideMouseOverIcon(0, 0, true);
	///------ End IMPROVE_HINT_IMAGE
	
	RECT rrList;
	GetWindowRect(rrList);
	
	int iTwipsWidth = PixelsToTwips( RECT_WIDTH(rrList) - 4 );
	int iTwipsHeight = PixelsToTwips( RECT_HEIGHT(rrList) );
	
	SetColWidth(0, iTwipsWidth);
	SetRowHeight(0, iTwipsHeight);	

	/// Iris 6/3/2015 ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
	//string strImageFile = GetFullPath(STR_PICT_NON_TEMPLATE_HINT ".emf", "Templates\\Previews", TRUE);
	string strName = lpcszImageFileName;
	strName += ".emf";
	string strImageFile = GetFullPath(strName, "Templates\\Previews", TRUE);
	///End ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
	/// Iris 6/11/2015 ORG-13203-P2 FIX_FAIL_TO_DISPLAY_NOPREVIEW_HINT_EMF_WHEN_TEMPALTE_NO_IMAGE_FILE
	//loadImage(strImageFile, 0, 0);
	loadImage(strImageFile, 0, 0, false);
	///End FIX_FAIL_TO_DISPLAY_NOPREVIEW_HINT_EMF_WHEN_TEMPALTE_NO_IMAGE_FILE
	return false;
}
///End IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE

///------ Tony 07/29/2015 ORG-13290-S6 ALLOW_SAVE_AS_TEMPLATE_FROM_TEMPLATE_LIBRARY
bool PreviewList::DoEdit(int nRow, int nCol)
{
	bool bRet = false;	
	if( IsCategoryRow(nRow) )
		bRet = editCategory(nRow);
	else
		bRet = editTemplate(nRow, nCol);

	return bRet;
}

bool PreviewList::editCategory(int nRow)
{
	ToggleCellEditControl(true, nRow, 0);
	return true;
}

bool PreviewList::editTemplate(int nRow, int nCol)
{
	string strOldPath;
	strOldPath = GetTemplateFilePath(nRow, nCol);
	GraphPage gp;
	if( !gp.Create( strOldPath ,CREATE_HIDDEN | CREATE_SET_MISSING_IN_MANAGER) || !gp )
		return error_report("fail to do edit since fail to create graph window with template.");

	string strOldName = GetFileName(strOldPath, true);
	
	int nXFRet = false;
	XFBase xfTemplateSaveAs("template_saveas");
	if( xfTemplateSaveAs )
	{
		///------ Tony 08/07/2015 ORG-13209-S5 NOT_ALLOW_SAME_TEMPLATE_NAME
		m_pXMLHelper->Save();
		///------ End NOT_ALLOW_SAME_TEMPLATE_NAME
		
		///------ Tony 08/04/2015 ORG-13290-S6 MODIFIED_TEMPLATE_NAME_IMPROVE
		Tree trRoot;
		TreeNode trGetN;
		xfTemplateSaveAs.GetGUI(trRoot, trGetN);
		trGetN.SetAttribute(STR_TITLE_ATTRIB, "template_modify");
		trGetN.SetAttribute(STR_COMMENT_ATTRIB, _L(":modify a graph template"));
		xfTemplateSaveAs.SetGUI(trGetN);
		///------ End MODIFIED_TEMPLATE_NAME_IMPROVE

		///------ Tony 10/14/2015 ORG-13290-S13 WANT_WARNING_WHEN_DELETING_TEMPLATE
		s_bOnActivateReady = false;
		///------ End WANT_WARNING_WHEN_DELETING_TEMPLATE
		
		Control* pctrl = GetControl();
		string strArguments;
		///------ Tony 08/14/2015 ORG-13290-P3 XF_ARGUMENT_SPACE_SUPPORT_IMPROVE
		//strArguments.Format("pg:=%s asksave:=2 emf:=0 bmp:=0 templateEdit:=%s", gp.GetName(), strOldName);
		strArguments.Format("pg:=%s asksave:=2 emf:=0 bmp:=0 templateEdit:=\"%s\"", gp.GetName(), strOldName);
		///------ End XF_ARGUMENT_SPACE_SUPPORT_IMPROVE
		///------ Folger 04/19/2018 ORG-18145-P1 FAIL_TO_MODIFY_OLD_FORMAT_TEMPLATE_IN_TEMPLATE_LIBRARY
		{
		LTVarTempChange OPJ("@OPJ", okutil_is_Origin_file_post_opj(strOldPath) == 0 ? 1 : 0);
		///------ End FAIL_TO_MODIFY_OLD_FORMAT_TEMPLATE_IN_TEMPLATE_LIBRARY
		nXFRet = xfTemplateSaveAs.ExecuteLabTalk(strArguments, NULL, NULL, XVT_INT, NULL, LTXF_SHOW_DIALOG | LTXF_FROM_CONTEXT_MENU, AU_NONE, pctrl ? pctrl->GetSafeHwnd() : NULL) == 0;
		///------ Folger 04/19/2018 ORG-18145-P1 FAIL_TO_MODIFY_OLD_FORMAT_TEMPLATE_IN_TEMPLATE_LIBRARY
		}
		///------ End FAIL_TO_MODIFY_OLD_FORMAT_TEMPLATE_IN_TEMPLATE_LIBRARY
	
		if(nXFRet)
		{
			Tree trRoot;
			TreeNode trGUI;
			xfTemplateSaveAs.GetGUI(trRoot, trGUI);
			string strNewName = trGUI.template.strVal;
			if(0 != strOldName.CompareNoCase(strNewName))
			{
				//new one added , delete old one
				DeleteFile(strOldPath);
				
				//rename emf and bmp file
				string strEmfNewPath = trGUI.filepath.strVal + strNewName + ".emf";
				string strEmfOldPath = trGUI.filepath.strVal + strOldName + ".emf";
				if(strEmfNewPath.IsFile())
					DeleteFile(strEmfNewPath);
				MoveFile(strEmfOldPath, strEmfNewPath);
				
				string strBmpNewPath = trGUI.filepath.strVal + strNewName + ".bmp";
				string strBmpOldPath = trGUI.filepath.strVal + strOldName + ".bmp";
				if(strBmpNewPath.IsFile())
					DeleteFile(strBmpNewPath);
				MoveFile(strBmpOldPath, strBmpNewPath);
			}
		}
		
		///------ Tony 08/07/2015 ORG-13209-S5 NOT_ALLOW_SAME_TEMPLATE_NAME
		Load(true);
		///------ End NOT_ALLOW_SAME_TEMPLATE_NAME
		s_bOnActivateReady = true;	///------ Tony 10/14/2015 ORG-13290-S13 WANT_WARNING_WHEN_DELETING_TEMPLATE
	}

	gp.Destroy();
	
	return nXFRet;
}
///------ End ALLOW_SAVE_AS_TEMPLATE_FROM_TEMPLATE_LIBRARY

/// Iris 3/4/2015 ORG-12495-S3 SUPPORT_DELETE_CATEGORY_AND_DELETE_TEMPLATE
bool PreviewList::Delete(int nRow, int nCol)
{
	///------ Tony 10/14/2015 ORG-13290-S13 WANT_WARNING_WHEN_DELETING_TEMPLATE
	string strWarning = _L("Are you sure you want to delete this item?");
	int nMessgeReturn = _MessageBox(s_hWndPreviewList, strWarning, _L("Attention!"), MB_YESNO | MB_DEFBUTTON2 | MB_ICONEXCLAMATION);
	if(IDNO == nMessgeReturn)
	{
		return false;
	}
	///------ End WANT_WARNING_WHEN_DELETING_TEMPLATE
	
	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	//if( IsCategoryRow(nRow) )
		//return deleteCategory(nRow);
	//else
		//return deleteTemplate(nRow, nCol);
	bool bRet = false;	
	if( IsCategoryRow(nRow) )
		bRet = deleteCategory(nRow);
	else
		bRet = deleteTemplate(nRow, nCol);
	
	CheckLoadImageForTempalteNone();
	
	return bRet;
	///------ End IMPROVE_HINT_IMAGE
}
	
bool PreviewList::deleteCategory(int nRow)
{
	int nCategoryIndex = GetCategoryIndex(nRow);
	if( nCategoryIndex < 0 )
		return false;
	
	int nLastImageRowInCategory = nRow + 1;
	while( !IsCategoryRow(nLastImageRowInCategory) && nLastImageRowInCategory + 1 < GetRows() )
		nLastImageRowInCategory++;
	
	if( IsCategoryRow(nLastImageRowInCategory) && nLastImageRowInCategory != nRow )
		nLastImageRowInCategory--;
	
	if( nLastImageRowInCategory - nRow + 1 > 0 )
	{
		for(int index = nLastImageRowInCategory; index >= nRow; index--)
			DeleteRow(index);		
		
		m_pXMLHelper->RemoveCategory(nCategoryIndex);	
	}
	return true;
}

bool PreviewList::deleteTemplate(int nRow, int nCol)
{

	///Sophy 7/27/2015 ORG-12495-P5 PROPER_UPDATE_CATEGORY_TEMPLATE_TREE_ON_MOVE_OR_DELETE_CELL moveImageToPreviousCell depends on the Category tree, so need to update after moveImageToPreviousCell 
	//// delete template from m_trUser
	//int nCategoryIndex = GetCategoryIndex(nRow);
	//int nTemplateIndex = GetTemplateIndex(nRow, nCol);					
	//m_pXMLHelper->RemoveTemplate(nCategoryIndex, nTemplateIndex, true, true);	
	///end PROPER_UPDATE_CATEGORY_TEMPLATE_TREE_ON_MOVE_OR_DELETE_CELL
	// delete image from list control
	int nLastImageCellRow, nLastImageCellCol;
	getLastImageCellInCategory(nRow, nCol, nLastImageCellRow, nLastImageCellCol);
	if( !( nRow == nLastImageCellRow && nCol == nLastImageCellCol ) )
		moveImageToPreviousCell(nRow, nCol, nLastImageCellRow, nLastImageCellCol);
	
	deleteImage(nLastImageCellRow, nLastImageCellCol, true);	
	
	///Sophy 7/27/2015 ORG-12495-P5 PROPER_UPDATE_CATEGORY_TEMPLATE_TREE_ON_MOVE_OR_DELETE_CELL
	// delete template from m_trUser
	int nCategoryIndex = GetCategoryIndex(nRow);
	///------ Tony 08/07/2015 ORG-13209-S5 NOT_ALLOW_SAME_TEMPLATE_NAME
	//int nTemplateIndex = GetTemplateIndex(nRow, nCol);		
	int nTemplateIndex = GetTemplateIndex(nRow, nCol, false);
	///------ End NOT_ALLOW_SAME_TEMPLATE_NAME
	m_pXMLHelper->RemoveTemplate(nCategoryIndex, nTemplateIndex, true, true);	
	///end PROPER_UPDATE_CATEGORY_TEMPLATE_TREE_ON_MOVE_OR_DELETE_CELL	
	
	return true;	
}
///End SUPPORT_DELETE_CATEGORY_AND_DELETE_TEMPLATE

// return true if added row
bool PreviewList::checkAddImageRow(int nRow)
{
	ASSERT(nRow >= 0 );
	if( nRow < 0 )
		return false;
		
	if( IsCategoryRow(nRow) )		
		m_flx.AddItem("", nRow);		
	else
	{
		if( nRow < GetNumRows() ) // to check first
			return false;
		
		while( nRow >= GetNumRows() ) 		
			AddOneRow();			
	}
	
	SetRowHeight(nRow, getCellWidth());	
	for(int nCol = 0; nCol < GetNumCols(); nCol++)
	{				
		SetColWidth(nCol, getCellWidth());
	}
	return true;
}

///------ Tony 07/23/2015 ORG-13290-S5 INITIALIZE_THE_CATEGORY_NAME_IMPROVE
//bool PreviewList::NewCategory()
bool PreviewList::NewCategory(LPCSTR strCategoryLabel)
///------ End INITIALIZE_THE_CATEGORY_NAME_IMPROVE
{	
	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	if( IsTemplateNone())
		return false;
	///------ End IMPROVE_HINT_IMAGE
	
	// add a empty new category to preivew list
	///------ Tony 07/14/2015 ORG-12495-S10 IMPROVE_TEMPLATE_FILTER
	//SetRows(GetRows() + 2); // 2, one for category text label, another for an empty row for image
	//
	//AddTextRow(GetRows() - 2, STR_NEW_CATEGORY);
	//
	//SetRowHeight(GetRows() - 1, getCellWidth());
	//SetColWidth(0, getCellWidth());	
	if(IsFilterOn())
	{
		SetRows(GetRows() + 1);
		///------ Tony 07/23/2015 ORG-13290-S5 INITIALIZE_THE_CATEGORY_NAME_IMPROVE
		//AddTextRow(GetRows() - 1, STR_NEW_CATEGORY);
		AddTextRow(GetRows() - 1, strCategoryLabel);
		///------ End INITIALIZE_THE_CATEGORY_NAME_IMPROVE
		SetColWidth(0, getCellWidth());	
	}
	else
	{
		SetRows(GetRows() + 2); // 2, one for category text label, another for an empty row for image
		///------ Tony 07/23/2015 ORG-13290-S5 INITIALIZE_THE_CATEGORY_NAME_IMPROVE
		//AddTextRow(GetRows() - 2, STR_NEW_CATEGORY);
		AddTextRow(GetRows() - 2, strCategoryLabel);
		///------ End INITIALIZE_THE_CATEGORY_NAME_IMPROVE
		SetRowHeight(GetRows() - 1, getCellWidth());
		SetColWidth(0, getCellWidth());	
	}
	
	///------ End IMPROVE_TEMPLATE_FILTER
	return true;
}

bool PreviewList::IsAllowDrop()
{
	/// Iris 4/13/2015 ORG-12495-S8 IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE
	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	//if( m_bIsTemplateNone )
	if( IsTemplateNone() )
	///------ End IMPROVE_HINT_IMAGE
		return false;
	///End IMPROVE_PREVIEW_FOR_NO_IMAGE_AND_NO_TEMPLATE
	
	int nDropRow, nDropCol;	
	
	///------ Tony 09/01/2015 ORG-12495-P6 DRAG_AND_DROP_IMPROVE
	//GetDropCell(nDropRow, nDropCol);
	GetDropCell(nDropRow, nDropCol, true);
	///------ End DRAG_AND_DROP_IMPROVE
	
	int nDraggedRow, nDraggedCol;
	///------ Tony 09/01/2015 ORG-12495-P6 DRAG_AND_DROP_IMPROVE
	//GetDraggedCell(nDraggedRow, nDraggedCol);
	GetDraggedCell(nDraggedRow, nDraggedCol, true);
	///------ End DRAG_AND_DROP_IMPROVE
	
	if( IsCategoryRow(nDraggedRow) != IsCategoryRow(nDropRow) )
		return false;
	
	if( !IsCategoryRow(nDraggedRow) ) // to drag image
	{
		ASSERT( HasImage(nDraggedRow, nDraggedCol) );
		
		if( !HasImage(nDropRow, nDropCol) ) // if the drop cell is empty ( no image )
		{
			int nPreviousImageCellRow, nPreviousImageCellCol;
			if( getPreviousImageCell(nDropRow, nDropCol, nPreviousImageCellRow, nPreviousImageCellCol) && !HasImage( nPreviousImageCellRow, nPreviousImageCellCol ) )				
				return false;			
			
			int nNextIamgeCellRow, nNextImageCellCol;
			if( getNextImageCell(nDraggedRow, nDraggedCol, nNextIamgeCellRow, nNextImageCellCol) && nNextIamgeCellRow == nDropRow && nNextImageCellCol == nDropCol )
				return false;
		}
	}
	else // to drag the whole category ??
	{
		
	}
	return true;
}

void PreviewList::DragAndDrop()
{
	if( !IsAllowDrop() )
		return;
	
	int nDraggedRow, nDraggedCol;
	///------ Tony 09/01/2015 ORG-12495-P6 DRAG_AND_DROP_IMPROVE
	//if( !GetDraggedCell(nDraggedRow, nDraggedCol) )
	if( !GetDraggedCell(nDraggedRow, nDraggedCol, true) )
	///------ End DRAG_AND_DROP_IMPROVE
		return;
	
	if( HasImage(nDraggedRow, nDraggedCol) )
		dragAndDropImage();
	else if( IsCategoryRow(nDraggedRow) )
		dragAndDropCategory();
}


void PreviewList::dragAndDropCategory()
{
	int nDraggedRow, nDraggedCol;
	int nDropRow, nDropCol;
	///------ Tony 09/01/2015 ORG-12495-P6 DRAG_AND_DROP_IMPROVE
	//if( !GetDraggedCell(nDraggedRow, nDraggedCol) )
	//	|| !GetDropCell(nDropRow, nDropCol) 
	if( !GetDraggedCell(nDraggedRow, nDraggedCol, true)
		|| !GetDropCell(nDropRow, nDropCol, true) 
	///------ End DRAG_AND_DROP_IMPROVE
		|| nDraggedRow == nDropRow && nDraggedCol == nDropCol )
	{		
		return;
	}
	
	if( !IsCategoryRow(nDraggedRow) || !IsCategoryRow(nDropRow) )
		return;
	
	int nDraggedCategory = GetCategoryIndex(nDraggedRow);
	int nDropCategory = GetCategoryIndex(nDropRow);
	
	if( nDraggedCategory < nDropCategory && nDraggedCategory + 1 == nDropCategory )		
		return;
	
	m_pXMLHelper->ReorderCategory(nDraggedCategory, nDropCategory);	
	Load();
}

void PreviewList::dragAndDropImage()
{	
	int nDraggedRow, nDraggedCol;
	int nDropRow, nDropCol;
	///------ Tony 09/01/2015 ORG-12495-P6 DRAG_AND_DROP_IMPROVE
	//if( !GetDraggedCell(nDraggedRow, nDraggedCol) )
	//	|| !GetDropCell(nDropRow, nDropCol) 
	if( !GetDraggedCell(nDraggedRow, nDraggedCol, true)
		|| !GetDropCell(nDropRow, nDropCol, true) 
	///------ End DRAG_AND_DROP_IMPROVE
		|| nDraggedRow == nDropRow && nDraggedCol == nDropCol )
	{		
		return;
	}	
	
	if( nDraggedRow == nDropRow )
	{
		if( nDropCol > 0 && nDropCol > nDraggedCol )
		{
			nDropCol--;		
		}
		else if( nDropCol < 0 )
		{
			nDropCol = GetNumCols() - 1;
		}
	}
	else
	{
		if( nDropCol < 0 )
		{
			nDropRow++;
			nDropCol = 0;
			
			if( IsCategoryRow(nDropRow) || nDropRow >= GetNumRows() )
			{
				checkAddImageRow(nDropRow);
			
				if( nDraggedRow >= nDropRow )
					nDraggedRow++;
			}
		}
	}	

	if( nDraggedRow == nDropRow && nDraggedCol == nDropCol 
		|| !HasImage(nDraggedRow, nDraggedCol)
		|| IsCategoryRow(nDropRow)
		|| nDropRow < 0 
		|| nDropCol < 0
		)
		return;	
	
	string strDraggedImageFilename = GetTemplateFilePath(nDraggedRow, nDraggedCol);	
	int nDraggedImageCategory = GetCategoryIndex(nDraggedRow);
	int nDraggedImageTemplate = GetTemplateIndex(nDraggedRow, nDraggedCol);
	
	int nDropCategory = GetCategoryIndex(nDropRow);
	int nDropTemplate = GetTemplateIndex(nDropRow, nDropCol, false);
	
	int nSuspendingTemplate = nDropTemplate;
	if ( NULL != m_pXMLHelper )
		m_pXMLHelper->AccessSuspendingTemplate(nDropCategory, nDropTemplate, AST_SET, strDraggedImageFilename);
	
	
	bool bRet;
	if( !HasImage(nDropRow, nDropCol) ) // if grag the current image to an empty cell
	{
		bRet = copyImage(nDraggedRow, nDraggedCol, nDropRow, nDropCol);	
		ASSERT(bRet);
		
		if ( nSuspendingTemplate < 0 && NULL != m_pXMLHelper )
		{
			nSuspendingTemplate = GetTemplateIndex(nDropRow, nDropCol);	
			m_pXMLHelper->AccessSuspendingTemplate(nDropCategory, nSuspendingTemplate, AST_SET, strDraggedImageFilename);
		}
			
		if( isLastImageCellInCategory(nDraggedRow, nDraggedCol) )
		{
			deleteImage(nDraggedRow, nDraggedCol, false);
		}
		else
		{	
			int nLastImageCellRow, nLastImageCellCol;
			getLastImageCellInCategory(nDraggedRow, nDraggedCol, nLastImageCellRow, nLastImageCellCol);
			
			moveImageToPreviousCell(nDraggedRow, nDraggedCol, nLastImageCellRow, nLastImageCellCol);		
			deleteImage(nLastImageCellRow, nLastImageCellCol, false);
		}
	}
	else
	{		
		// to check if drag & drop image in one category
		int nDraggedCategoryLastImageCellRow, nDraggedCategoryLastImageCellCol;
		int nDropCategoryLastImageCellRow, nDropCategoryLastImageCellCol;
		getLastImageCellInCategory(nDraggedRow, nDraggedCol, nDraggedCategoryLastImageCellRow, nDraggedCategoryLastImageCellCol);
		getLastImageCellInCategory(nDropRow, nDropCol, nDropCategoryLastImageCellRow, nDropCategoryLastImageCellCol);
		bool bDragDropInOneCategory = ( nDraggedCategoryLastImageCellRow == nDropCategoryLastImageCellRow && nDraggedCategoryLastImageCellCol == nDropCategoryLastImageCellCol );
		
		// reorder the related images
		if( bDragDropInOneCategory )
		{
			// if drag the current image to previous cells in same category
			if( nDropRow < nDraggedRow || nDropRow == nDraggedRow && nDropCol < nDraggedCol )
				moveImageToNextCell(nDraggedRow, nDraggedCol, nDropRow, nDropCol);
			else if( nDropRow > nDraggedRow || nDropRow == nDraggedRow && nDropCol > nDraggedCol )// if drag the current image to back cell in same category
				moveImageToPreviousCell(nDraggedRow, nDraggedCol, nDropRow, nDropCol);
			
		}
		else
		{
			// 1. For current category
			// arrange the remaining cells in category of the dragged cell
			int nNextRow, nNextCol;
			int nRow = nDraggedRow, nCol = nDraggedCol;
			int nCopyImageCount;
			while( getNextImageCell(nRow, nCol, nNextRow, nNextCol, false) && HasImage(nNextRow, nNextCol) )
			{				
				copyImage(nNextRow, nNextCol, nRow, nCol);				
				
				nRow = nNextRow;
				nCol = nNextCol;
				nCopyImageCount++;
			}
			deleteImage(nRow, nCol, false); // to delete the last image cell in this category
			
			// if there is not any image cell at the next of the dragged image. 
			// Just need to delete the dragged image from the current category. 
			if( 0 == nCopyImageCount )			
				deleteImage(nDraggedRow, nDraggedCol, false);
			
			// 2. For the new category
			bool bAddedNewRow;
			int nNextOfCurrentLastRow, nNextOfCurrentLastCol;
			if( getNextImageCell(nDropCategoryLastImageCellRow, nDropCategoryLastImageCellCol, nNextOfCurrentLastRow, nNextOfCurrentLastCol, true, &bAddedNewRow) )
			{
				if( bAddedNewRow && nDraggedRow > nDropRow )
					nDraggedRow++;
				moveImageToNextCell(nNextOfCurrentLastRow, nNextOfCurrentLastCol, nDropRow, nDropCol);
			}				
		}
		
		// put the current image to drop row * col					
		loadImage(strDraggedImageFilename, nDropRow, nDropCol);		
	}	
	
	if ( NULL != m_pXMLHelper )
		m_pXMLHelper->AccessSuspendingTemplate(nDropCategory, nSuspendingTemplate, AST_DELETE);
	
	
    bool bInOneCategory = (nDraggedImageCategory == nDropCategory);
	
	bool bUntitledCategoryRemoved;
	m_pXMLHelper->RemoveTemplate(nDraggedImageCategory, nDraggedImageTemplate, false, !bInOneCategory, &bUntitledCategoryRemoved);	
	if( bUntitledCategoryRemoved )
		nDropCategory--;
	
	if( bInOneCategory && nDropTemplate > nDraggedImageTemplate )
		nDropTemplate++;
	m_pXMLHelper->InsertTemplate(nDropCategory, nDropTemplate, strDraggedImageFilename, true);

	// remove the empty image row	
	int nLastCellRow, nLastCellCol;
	getLastImageCellInCategory(nDraggedRow, nDraggedCol, nLastCellRow, nLastCellCol);
			
	// remove the empty image row
	if( HasImage(nLastCellRow, nLastCellCol) && nLastCellCol == GetNumCols() - 1 )
		nLastCellRow++;
	
	if( !IsCategoryRow(nLastCellRow) )
	{
		int nEmptyCount;
		for(int nCol = 0; nCol < GetNumCols(); nCol++)
		{
			if( !HasImage(nLastCellRow, nCol) )
				nEmptyCount++;
		}
		if( nEmptyCount == GetNumCols() )
		{			
			DeleteRow(nLastCellRow);
			if( nLastCellRow < nDropRow )
				nDropRow--; // adjust row index after delete one row
		}
	}
	
	SelCell(nDropRow, nDropCol); // to refresh when drop image to an empty cell
}

bool PreviewList::getNextImageCell(int nRow, int nCol, int& nNextCellRow, int& nNextCellCol, bool bAddIfNotFound, bool* pbIsAddedRow)
{
	if( nCol == GetNumCols() - 1 ) // last col
	{
		// goto the next row 
		int nNextRow = nRow + 1;	
		if( nNextRow == GetNumRows() || IsCategoryRow(nNextRow)  )
		{
			if( !bAddIfNotFound )
				return false;		
			
			bool bRet = checkAddImageRow(nNextRow);
			
			if( NULL != pbIsAddedRow )
				*pbIsAddedRow = bRet;
		}	
		
		nNextCellRow = nNextRow;
		nNextCellCol = 0;		
		return true;
	}
	else if( nCol < GetNumCols() -1 )
	{
		nNextCellRow = nRow;
		nNextCellCol = nCol + 1;
		return true;
	}
	return false;
}

bool PreviewList::getPreviousImageCell(int nRow, int nCol, int& nPreviousCellRow, int& nPreviousCellCol)
{
	if( 0 == nRow && 0 == nCol )
		return false;
	
	bool bRet;
	if( 0 == nCol )
	{
		if( nRow - 1 >= 0 )
		{
			nPreviousCellRow = nRow - 1;
			nPreviousCellCol = GetNumCols() - 1;
			bRet = true;
		}
	}
	else 
	{
		nPreviousCellRow = nRow;
		nPreviousCellCol = nCol - 1;
		bRet = true;
	}
	
	return bRet && !IsCategoryRow(nPreviousCellRow);
}

// to check the image cell if is the last one in category.
bool PreviewList::isLastImageCellInCategory(int nRow, int nCol)
{
	if( nCol < GetNumCols() - 1 )
		return !HasImage(nRow, nCol + 1);
	else
		return IsCategoryRow(nRow + 1);
}

void PreviewList::getLastImageCellInCategory(int nRow, int nCol, int& nLastCellRow, int& nLastCellCol)
{
	if( !HasImage(nRow, nCol) )
	{
		int nPreviousCellRow, nPreviousCellCol;	
		while( getPreviousImageCell(nRow, nCol, nPreviousCellRow, nPreviousCellCol) )
		{
			if( HasImage(nPreviousCellRow, nPreviousCellCol) )
			{
				nLastCellRow = nPreviousCellRow;
				nLastCellCol = nPreviousCellCol;
				break;
			}
		}
	}
	else
	{
		nLastCellRow = nRow;
		nLastCellCol = nCol;
		
		int nNextCellRow, nNextCellCol;	
		while( getNextImageCell(nRow, nCol, nNextCellRow, nNextCellCol) )
		{
			if( !HasImage(nNextCellRow, nNextCellCol) )
				break;
			
			nLastCellRow = nNextCellRow;
			nLastCellCol = nNextCellCol;
			
			if( IsCategoryRow(nNextCellRow) )
				break;		
			
			nRow = nNextCellRow;
			nCol = nNextCellCol;
		}	
	}
}

bool TemplateListBase::HasImage(int nRow, int nCol)
{
	PictureHolder pict;
	bool bRet = GetCellPicture(nRow, nCol, pict);
	return (bRet && NULL != pict);
}

bool PreviewList::getImage(int nRow, int nCol, PictureHolder& pict, string* pstrFilename)
{
	if ( GetCellPicture(nRow, nCol, pict) && NULL != pict )
	{
		if( pstrFilename )
			*pstrFilename = GetTemplateFilePath(nRow, nCol);		
		return true;
	}
	return false;
}

bool PreviewList::setImage(int nRow, int nCol, PictureHolder pict, LPCSTR lpcszFilename, bool bShowFileName)
{
	if( nRow >= GetRowOffset() && nCol >= GetColOffset() && nRow < GetNumRows() && nCol < GetNumCols() )
	{
		SetCellPicture(nRow, nCol, pict);
		SetCell(flexPicAlignCenterCenter, flexcpPictureAlignment, nRow, nCol);
		
		///------ Tony 11/23/2016 ORG-15717-P1 DELELE_IMAGE_SET_DATA_NULL
		if( NULL == lpcszFilename)
		{
			SetCell(lpcszFilename, flexcpData, nRow, nCol);
			SetCell(lpcszFilename, flexcpText, nRow, nCol);
		}
		else
		///------ End DELELE_IMAGE_SET_DATA_NULL
		
		if( NULL != lpcszFilename && bShowFileName )
		{		
			SetCell(lpcszFilename, flexcpData, nRow, nCol);
			SetCell(GetFileName(lpcszFilename, true), flexcpText, nRow, nCol);
			SetCell(flexAlignCenterTop, flexcpAlignment, nRow, nCol);
			if ( pict.GetWidth() * 0.9 <= pict.GetHeight() )
				SetCell(80, flexcpPictureScaleFactor, nRow, nCol);
		}
		
		bool bMatch = _is_template_match_with_active_sheet(lpcszFilename);
		bool bCloneable = _is_template_cloneable(lpcszFilename);
		int nPosition = bCloneable ? (bMatch ? flexIndicatorPositionLeftBottomHighlight : flexIndicatorPositionLeftBottom) : flexIndicatorPositionNone;
		
		SetCell(nPosition, flexcpPictureIndicatorPosition, nRow, nCol);
		
		return true;
	}
	return false;
}

bool PreviewList::copyImage(int nRowFrom, int nColFrom, int nRowTo, int nColTo)
{
	PictureHolder pict;
	string strFilename;
	if( !getImage(nRowFrom, nColFrom, pict, &strFilename) )
		return false;
	
	return setImage(nRowTo, nColTo, pict, strFilename);
}

bool PreviewList::deleteImage(int nRow, int nCol, bool bDeleteRow)
{	
	if( !HasImage(nRow, nCol) )	
		return false;
	
	PictureHolder pictNULL;
	
	setImage(nRow, nCol, pictNULL, NULL);
	
	if( bDeleteRow && 0 == nCol && /*0 != nRow && */!IsCategoryRow(nRow - 1) ) // delete the empty row if the the row is not the last 
		DeleteRow(nRow);
	return true;
}

void PreviewList::moveImageToNextCell(int nBeginRow, int nBeginCol, int nEndRow, int nEndCol)
{		
	int nRow = nBeginRow, nCol = nBeginCol;
	while( true )
	{
		int nPreviousRow, nPreviousCol;
		getPreviousImageCell(nRow, nCol, nPreviousRow, nPreviousCol);
		
		PictureHolder pict;
		string strFilename;
		if( !getImage(nPreviousRow, nPreviousCol, pict, &strFilename) )
			break;
		
		setImage(nRow, nCol, pict, strFilename);
		
		if( nPreviousRow == nEndRow && nPreviousCol == nEndCol )
			break;
		nRow = nPreviousRow;
		nCol = nPreviousCol;
	}	
}

void PreviewList::moveImageToPreviousCell(int nBeginRow, int nBeginCol, int nEndRow, int nEndCol)
{
	int nRow = nBeginRow, nCol = nBeginCol;
	while( true )
	{
		int nNextRow, nNextCol;
		if( !getNextImageCell(nRow, nCol, nNextRow, nNextCol, false) )	
			break;
		
		PictureHolder pict;
		string strFilename;
		if( !getImage(nNextRow, nNextCol, pict, &strFilename) )
			break;
		
		setImage(nRow, nCol, pict, strFilename);		
		
		if( nNextRow == nEndRow && nNextCol == nEndCol )
			break;
		nRow = nNextRow;
		nCol = nNextCol;		
	}	
}

int PreviewList::GetCurrentCategory()
{
	int nx, ny, row, col;
	GetSelCell(nx, ny, row, col);
	if( row < 0 || col < 0 )
		return -1;
	
	return GetCategoryIndex(row);
}

int PreviewList::GetCategoryIndex(int nRow)
{	
	/// Iris 6/3/2015 ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
	if( 0 == m_pXMLHelper->GetCategoryCount() )
		return -1;
	///End ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
	
	/// Iris 6/3/2015 ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
	int row = nRow;
	while( !IsCategoryRow(row) && row > GetRowOffset() )
		row--;
	///------ Tony 07/14/2015 ORG-12495-S10 IMPROVE_TEMPLATE_FILTER
	//if( row <= GetRowOffset() )
	if( row <= GetRowOffset() && !IsCategoryRow(row) )
	///------ End IMPROVE_TEMPLATE_FILTER
		return NUM_UNTITLED_CATEGORY_ID;
	///End ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
	
	if( IsFilterOn() && m_trFilter.FirstNode )
	{
		int nCategoryRow = nRow;
		while( !IsCategoryRow(nCategoryRow) &&nCategoryRow >= GetRowOffset() )
			nCategoryRow--;		
		
		string strCategoryLabel = GetCell(nCategoryRow, 0);
		ASSERT( !strCategoryLabel.IsEmpty() );
		
		TreeNode trCategory = m_trFilter.FindNodeByAttribute(STR_LABEL_ATTRIB, strCategoryLabel, false);		
		
		int nDataID;
		bool bRet = trCategory && trCategory.GetAttribute(STR_DATAID_ATTRIB, nDataID);
		ASSERT( bRet );
		
		return nDataID;
	}

	return TemplateListBase::GetCategoryIndex(nRow);	
}

int PreviewList::GetTemplateIndex(int nRow, int nCol, bool bCheckImage /*= true*/)
{
	if( bCheckImage && !HasImage(nRow, nCol) )
        return -1;
	
	
	// look forward what the number of templates at the current template.
	int nCurrentTemplateIndex = 0;
	
	int nPreviousRow, nPreviousCol;
	int row = nRow, col = nCol;
	while( getPreviousImageCell(row, col, nPreviousRow, nPreviousCol) )
	{
		nCurrentTemplateIndex++;
		
		row = nPreviousRow;
		col = nPreviousCol;
	}
	
	if( IsFilterOn() && m_trFilter.FirstNode )
	{
		// find category by category label text in m_trFilter.
		while( !IsCategoryRow(nRow) && nRow > GetRowOffset() )
			nRow--;
		
		TreeNode trCategory;
		if( IsCategoryRow(nRow) )
		{
			string strCategoryLabel = GetCell(nRow, 0);
			trCategory = m_trFilter.FindNodeByAttribute(STR_LABEL_ATTRIB, strCategoryLabel, false);
		}
		else
			trCategory = m_trFilter.FindNodeByAttribute(STR_DATAID_ATTRIB, NUM_UNTITLED_CATEGORY_ID, false);			
		ASSERT( trCategory && nCurrentTemplateIndex < trCategory.Children.Count() );
		
		if( trCategory )
		{
			TreeNode trTemplate = trCategory.Children.Item(nCurrentTemplateIndex);
			ASSERT( trTemplate );
			
			int index;
			if( trTemplate && trTemplate.GetAttribute(STR_DATAID_ATTRIB, index) )
				return index;		
		}
	}
	
	return nCurrentTemplateIndex;
}

string PreviewList::GetTemplateFilePath(int nRow, int nCol)
{
	if( IsCategoryRow( nRow ) )
		return "";
	
	int nCategoryIndex = GetCategoryIndex(nRow);		
	int nTemplateIndex = GetTemplateIndex(nRow, nCol);
	
	if( nCategoryIndex < 0 || nTemplateIndex < 0 )
		return "";
	
	
	string strFilename;		
	bool bRet = m_pXMLHelper->GetTemplateFilePath(nCategoryIndex, nTemplateIndex, strFilename);			
	//ASSERT(bRet);
	return strFilename;
}

int PreviewList::getCellWidth()
{
	return NUM_TEMPLATE_PREVIEW_WIDTH;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////  TemplateSettingsHelper  /////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////

bool TemplateSettingsHelper::Init(int nPageType)
{
	m_nPageType = nPageType;
	m_bModified = false;
	///------ Tony 08/07/2015 ORG-13209-S5 NOT_ALLOW_SAME_TEMPLATE_NAME
	//return loadXML();
	return LoadXML();
	///------ End NOT_ALLOW_SAME_TEMPLATE_NAME
}

static	TreeNode	_add_category(TreeNode& trUser, string strCategoryLabel)
{
	string strCateNode = tree_get_enum_node_name(trUser, STR_CATEGORY_NODE);
	TreeNode trCate = trUser.AddNode(strCateNode, TRGP_BRANCH);
	if ( strCategoryLabel.IsEmpty() )
		strCategoryLabel = tree_get_enum_attribute_value(trUser, STR_LABEL_ATTRIB, STR_NEW_CATEGORY_PREFIX);
	
	trCate.SetAttribute(STR_USER_PATH_ATTRIB, 1);
	trCate.SetAttribute(STR_LABEL_ATTRIB, strCategoryLabel);
	return trCate;
	
}
static	void	_add_one_xml_to_tree(TreeNode& trUser, int nPageType, int nLocation)
{
	Tree tr;
	if( load_template_tree(tr, nLocation) )
	{
		TreeNode trGraphBranch = tree_get_node_by_dataid(tr, nPageType);
		if( trGraphBranch )
		{
			foreach(TreeNode trCate in trGraphBranch.Children)
			{
				string strLabel;
				trCate.GetAttribute(STR_LABEL_ATTRIB, strLabel);
				TreeNode trCategory  = trUser.FindNodeByAttribute(STR_LABEL_ATTRIB, strLabel);
				if ( !trCategory )
				{										
					trCategory = _add_category(trUser, strLabel);
				}
				foreach(TreeNode trOne in trCate.Children)
				{
					string strTemplate, strFile;
					trOne.GetAttribute(STR_LABEL_ATTRIB, strTemplate);
					trOne.GetAttribute(STR_FILENAME_ATTRIB, strFile);
					
					TreeNode trNode;
					okutil_add_template_node(&trNode, &trCategory, strTemplate, strFile, false);
				}
			}
		}
	}	
}

bool TemplateSettingsHelper::LoadXML()
{	
	m_trUser.Reset();	
	
	_add_one_xml_to_tree(m_trUser, m_nPageType, TEMPLATE_LOCATION_USER_FOLDER);
	_add_one_xml_to_tree(m_trUser, m_nPageType, TEMPLATE_LOCATION_GROUP_FOLDER);
	foreach(TreeNode trCat in m_trUser.Children)
	{
		setupTemplateNodeID(trCat);
	}
	setupCategoryNodeID(m_trUser);
	
	return true;
}


TreeNode TemplateSettingsHelper::checkGetUntitledCategory()
{
	string strName = STR_UNTITLED_CATEGORY_NAME;
	TreeNode trUntitledCategory = m_trUser.GetNode(strName);
	if( !trUntitledCategory )
	{		
		if( m_trUser.FirstNode )
			trUntitledCategory = m_trUser.InsertNode(m_trUser.FirstNode, strName);
		else
			trUntitledCategory = m_trUser.AddNode(strName);
	
		trUntitledCategory.SetAttribute(STR_DATAID_ATTRIB, NUM_UNTITLED_CATEGORY_ID);
		trUntitledCategory.SetAttribute(STR_LABEL_ATTRIB, STR_UNTITLED_CATEGORY_LABEL);		
	}
	return trUntitledCategory;
}

bool TemplateSettingsHelper::ScanUserDefinedTemplate()
{
	StringArray saTemplateFiles;
	FindFiles(saTemplateFiles, STR_ORIGIN_USER_FILE_FOLDER, GetTemplateExt());
	if( 0 == saTemplateFiles.GetSize() )
		return false;	
	
	TreeNode trUntitledCategoryNode = checkGetUntitledCategory();	
	
	int count = 0;
	for(int ii = 0; ii < saTemplateFiles.GetSize(); ii++)
	{
		if( 0 == saTemplateFiles[ii].CompareNoCase("layout.otp") )
			continue;
		
		if( !CheckCategory(STR_ORIGIN_USER_FILE_FOLDER + saTemplateFiles[ii]) )
		{			
			if(!CheckSetTemplateName(STR_ORIGIN_USER_FILE_FOLDER + saTemplateFiles[ii], NUM_UNTITLED_CATEGORY_ID, &count)) ///------ Tony 08/12/2015 ORG-13290-S8 ADD_SCAN_TEMPLATE_DETECT_DUPLICATE_NAME
			{
				// append template node
				InsertTemplate(NUM_UNTITLED_CATEGORY_ID, -1, STR_ORIGIN_USER_FILE_FOLDER + saTemplateFiles[ii], false);
			
				count++;
			}
		}
		
	}
	
	if( count > 0 )
		setupTemplateNodeID(trUntitledCategoryNode);	
	else if( !trUntitledCategoryNode.FirstNode )
		trUntitledCategoryNode.Remove();	
	
	return ( count > 0 );
}

bool TemplateSettingsHelper::CheckSetTemplateName(LPCSTR lpcszTemplateFile, int nCategory, int* pnCount)
{
	string strSrcFileName = lpcszTemplateFile;
	string strSrcName = GetFileName(strSrcFileName, true);
	foreach(TreeNode trCategory in m_trUser.Children)
	{
		foreach(TreeNode trTemplate in trCategory.Children)
		{
			string strFileName;
			trTemplate.GetAttribute(STR_FILENAME_ATTRIB, strFileName);
			string strName = GetFileName(strFileName, true);
			if(0 == strName.CompareNoCase(strSrcName))
			{
				string strWarning;
				ocu_load_msg_str(IDS_TEMPLATE_SAME_NAME_IN_LIBRARY_FOR_TL, &strWarning, "\"" + strName + "\"");
				
				int nRet = _MessageBox(s_hWndPreviewList, strWarning, _L("Warning"), MB_YESNOCANCEL|MB_ICONEXCLAMATION);
				
				//Yes will show the new template instead.
				if(IDYES == nRet)
				{
					trTemplate.Remove();
						
					InsertTemplate(nCategory, -1, strSrcFileName, false);
			
					if(pnCount)
						(*pnCount)++;
					
					//if in user path, then the old template should deleted
					string strFolder = GetFilePath(strFileName);
					if(0 == strFolder.CompareNoCase(STR_USER_PATH))
					{
						string strDeleteFolder = STR_USER_PATH + "DeletedTemplates\\";
						CheckMakePath(strDeleteFolder);
						
						string strDeletePath = strDeleteFolder + GetFileName(strFileName, false);
						
						if( strDeletePath.IsFile() )
							DeleteFile(strDeletePath);	
							
                        MoveFile(strFileName, strDeletePath);
						
						vector<string> vsImageExt = {".EMF", ".BMP"};
						for(int ii = 0; ii < vsImageExt.GetSize(); ii++)
						{
							string strImageFile = GetFilePath(strFileName) + GetFileName(strFileName, true) + vsImageExt[ii];						
							if( strImageFile.IsFile() )
							{
								string strDeleteImagePath = strDeleteFolder + GetFileName(strFileName, true) + vsImageExt[ii];

								if( strDeleteImagePath.IsFile() )
									DeleteFile(strDeleteImagePath);					
								MoveFile(strImageFile, strDeleteImagePath);
							}
						}				
					}
				}
				//No will auto rename the new template.
				else if(IDNO == nRet)
				{
					string strNewName = strSrcName;
					
					vector<string> vsAttribVals;
					foreach(TreeNode trCategoryTemp in m_trUser.Children)
					{
						vector<string> vsOneAttribVals;
						tree_get_attributes(trCategoryTemp, vsOneAttribVals, STR_LABEL_ATTRIB, true);
						vsAttribVals.Append(vsOneAttribVals);
					}
					
					int nNum = get_list_enum_number(vsAttribVals, strNewName);
					string strNum = 0 == nNum ? "" : (string)nNum;
					strNewName = strNewName + strNum;
					
					string strEtc = get_file_ext(strSrcFileName);
					string strNewFileName = GetFilePath(strSrcFileName) + strNewName + strEtc;
					
					if( strSrcFileName.IsFile() )
						RenameFile(strNewFileName, strSrcFileName);
					// EMF file
					vector<string> vsImageExt = {".EMF", ".BMP"};
					for(int ii = 0; ii < vsImageExt.GetSize(); ii++)
					{
						string strImageFile = GetFilePath(strSrcFileName) + GetFileName(strSrcFileName, true) + vsImageExt[ii];						
						if( strImageFile.IsFile() )
						{
							string strNewImageFile = GetFilePath(strSrcFileName) + strNewName + vsImageExt[ii];	
							RenameFile(strNewImageFile, strImageFile);
						}
					}		
					
					// append template node
					InsertTemplate(nCategory, -1, strNewFileName, false);
			
					if(pnCount)
						(*pnCount)++;
				}
				//Cancel will not add this template.
				else
				{

				}
				
				return true;
			}
		}
	}
	
	
	return false;
}

bool TemplateSettingsHelper::CheckCategory(LPCSTR lpcszTemplateFile, int* pnCategoryID, string* pstrCategoryName)
{
	if( m_trUser.FirstNode )
	{
		TreeNode trN = m_trUser.FindNodeByAttribute(STR_FILENAME_ATTRIB, lpcszTemplateFile, true);
		if( trN ) // means in one category
		{
			TreeNode trCategory = trN.Parent();	
			
			if( trCategory )
			{
				int nCategoryID;
				if( trCategory.GetAttribute(STR_DATAID_ATTRIB, nCategoryID) && pnCategoryID )
					*pnCategoryID = nCategoryID;				
				
				string str;
				if( pstrCategoryName && trCategory.GetAttribute(STR_LABEL_ATTRIB, str) )
					*pstrCategoryName = str;		
				
				return true;
			}
		}		
	}	
	return false;
}

void TemplateSettingsHelper::ReorderCategory(int nCurrent, int nNew)
{
	TreeNode trCurrent = getCategoryNode(nCurrent);
	TreeNode trNew = getCategoryNode(nNew);
	if( trCurrent )
	{
		if( !trNew )
			m_trUser.AddNode(trCurrent);
		else
		{
			TreeNode trN = m_trUser.InsertNode(trNew, "junk");
			trN.Replace(trCurrent);
		}
		trCurrent.Remove();
		setupCategoryNodeID(m_trUser);
	}
}

/// Iris 3/4/2015 ORG-12495-S3 SUPPORT_DELETE_CATEGORY_AND_DELETE_TEMPLATE
bool	TemplateSettingsHelper::RemoveCategory(int index)
{
	TreeNode trNode = getCategoryNode(index);
	if( trNode )
	{
		trNode.Remove();
		setupCategoryNodeID(m_trUser);
		SetModifyFlag();
		return true;
	}
	return false;
}
///End SUPPORT_DELETE_CATEGORY_AND_DELETE_TEMPLATE

TreeNode TemplateSettingsHelper::getCategoryNode(int index)
{	
	if( m_trUser.FirstNode && index >= 0 )
	{	
		///------ Tony 07/23/2015 ORG-13290-S5 INITIALIZE_THE_CATEGORY_NAME_IMPROVE
		//TreeNodeCollection trCategoryColl(m_trUser, "Category");
		//if( index < trCategoryColl.Count() )
		//{
			//TreeNode trCategory = trCategoryColl.Item(index);	
			//if( trCategory )
				//return trCategory;	
		//}
		//can not use TreeNodeCollection because it will sort automatically
		if( index < m_trUser.Children.Count() )
		{
			return m_trUser.Children.Item(index);
		}
		///------ End INITIALIZE_THE_CATEGORY_NAME_IMPROVE
	}
	
	TreeNode trJunk;	
	return trJunk;	
}

TreeNode TemplateSettingsHelper::getTemplateNode(TreeNode& trCategory, int index)
{
	return trCategory.FindNodeByAttribute(STR_DATAID_ATTRIB, index + 1, false);
}

bool TemplateSettingsHelper::GetCategoryLabel(int index, string& strLabel)
{
	TreeNode trCategory = getCategoryNode(index);
	return trCategory && trCategory.GetAttribute(STR_LABEL_ATTRIB, strLabel);	
}

///Sophy 7/27/2015 ORG-12495-P5 PROPER_UPDATE_CATEGORY_TEMPLATE_TREE_ON_MOVE_OR_DELETE_CELL
void	TemplateSettingsHelper::AccessSuspendingTemplate(int nCategory, int nTempalteIndex, int nOpt, string& strTemplateFileName)
{
	if ( nCategory < 0 || nTempalteIndex < 0 )
		return;
	string	strKey;
	strKey.Format("Catetory%d_Template%d", nCategory, nTempalteIndex);
	if ( AST_DELETE == nOpt )
	{
		m_trSuspending.RemoveChild(strKey);
	}
	else if ( AST_SET == nOpt )
	{
		TreeNode trNode = tree_check_get_node(m_trSuspending, strKey);
		if ( trNode )
			trNode.strVal = strTemplateFileName;
	}
	else if ( AST_GET == nOpt && NULL != strTemplateFileName )
	{
		TreeNode trNode = m_trSuspending.GetNode(strKey);
		if ( trNode )
			strTemplateFileName = trNode.strVal;
	}
}
///endPROPER_UPDATE_CATEGORY_TEMPLATE_TREE_ON_MOVE_OR_DELETE_CELL

bool TemplateSettingsHelper::GetTemplateFilePath(int nCategoryIndex, int nTemplateIndex, string& strFilename)
{	
	TreeNode trCategory = getCategoryNode(nCategoryIndex);
	if( trCategory && trCategory.FirstNode  )
	{		
		if( nTemplateIndex < trCategory.Children.Count() )
		{
			TreeNode trTemplate = trCategory.Children.Item(nTemplateIndex);
			if( trTemplate )
				return trTemplate.GetAttribute(STR_FILENAME_ATTRIB, strFilename);
		}
		///Sophy 7/27/2015 ORG-12495-P5 PROPER_UPDATE_CATEGORY_TEMPLATE_TREE_ON_MOVE_OR_DELETE_CELL
		else
		{
			AccessSuspendingTemplate(nCategoryIndex, nTemplateIndex, AST_GET, strFilename);
			return !strFilename.IsEmpty();
		}
		///end PROPER_UPDATE_CATEGORY_TEMPLATE_TREE_ON_MOVE_OR_DELETE_CELL
	}
	return false;
}

// once add/remove or change the position of template in this category, need to call the function to reset ID for all template nodes.
static void _tree_reset_node_data_id(TreeNode& trBranch, LPCSTR lpcszSkipNodeLabel = NULL)
{
	if( !trBranch )
		return;
	
	int id = 0;
	foreach(TreeNode trNode in trBranch.Children)
	{
		string strLabel;
		trNode.GetAttribute(STR_LABEL_ATTRIB, strLabel);
		
		if( NULL == lpcszSkipNodeLabel || 0 != strLabel.CompareNoCase(lpcszSkipNodeLabel) )
			trNode.SetAttribute(STR_DATAID_ATTRIB, ++id); // id begin from 1
	}
}

static void _remove_template_label_attribue(TreeNode &tr)
{	
	foreach(TreeNode trCategory in tr.Children)	
		octree_remove_attribute(&trCategory, STR_LABEL_ATTRIB);		
}

void TemplateSettingsHelper::setupCategoryNodeID(TreeNode& tr)
{	
	_tree_reset_node_data_id(tr, STR_UNTITLED_CATEGORY_LABEL);
	
	TreeNode trN = m_trUser.FindNodeByAttribute(STR_LABEL_ATTRIB, STR_UNTITLED_CATEGORY_LABEL);
	if( trN )
		trN.SetAttribute(STR_DATAID_ATTRIB, NUM_UNTITLED_CATEGORY_ID);
}

void TemplateSettingsHelper::setupTemplateNodeID(TreeNode& trCategory)
{
	_tree_reset_node_data_id(trCategory); // begin from 1
}

bool TemplateSettingsHelper::NewCategory(string& strCategoryLabel)
{
	string strCategoryName = tree_get_enum_node_name(m_trUser, STR_CATEGORY_NODE);
	TreeNode trNewCategory = tree_check_get_node(m_trUser, strCategoryName);

	strCategoryLabel = tree_get_enum_attribute_value(m_trUser, STR_LABEL_ATTRIB, STR_NEW_CATEGORY);
	trNewCategory.SetAttribute(STR_LABEL_ATTRIB, strCategoryLabel);

	
	setupCategoryNodeID(m_trUser);
	
	SetModifyFlag();
	return true;
}

bool TemplateSettingsHelper::RenameCategory(int index, string& strNewName)
{
	TreeNode trCategory = getCategoryNode(index);
	if(!trCategory)
		return false;
	
	if(strNewName.IsEmpty() || 0 == strNewName.CompareNoCase(STR_UNTITLED_CATEGORY_LABEL))
	{
		string strOldName;
		trCategory.GetAttribute(STR_LABEL_ATTRIB, strOldName);
		strNewName = strOldName;
	}
	else
	{
		//remove the current treenode
		TreeNode trUser = m_trUser.Clone();
		trUser.RemoveChild(trCategory.tagName);
		
		vector<string> vsAttribVals;
		tree_get_attributes(trUser, vsAttribVals, STR_LABEL_ATTRIB, true);
		if ( -1 != vsAttribVals.Find(strNewName) )
		{
			strNewName = tree_get_enum_attribute_value(trUser, STR_LABEL_ATTRIB, strNewName);
		}
		trCategory.SetAttribute(STR_LABEL_ATTRIB, strNewName);
		
		SetModifyFlag();
	}
	
	return true;
}

void TemplateSettingsHelper::SetModifyFlag()
{
	m_bModified = true;
}

bool TemplateSettingsHelper::IsModified()
{
	return m_bModified;
}

bool TemplateSettingsHelper::Save()
{
	if( !IsModified() )
		return false;	
	
	_remove_template_label_attribue(m_trUser);
	tree_remove_attribute(m_trUser, STR_DATAID_ATTRIB);
	
	Tree trXML;
	string strUserXML = STR_TEMPLATE_XML;
	load_template_tree(trXML, !strUserXML.IsFile() ? TEMPLATE_LOCATION_SYSTEM_FOLDER : TEMPLATE_LOCATION_USER_FOLDER);
	
	TreeNode trGraphBranch = tree_get_node_by_dataid(trXML, m_nPageType);
	trGraphBranch.Replace(m_trUser, true, true, true); 
	
	trXML.Save(strUserXML);
	return true;
}

int TemplateSettingsHelper::GetCategoryCount()
{
	TreeNodeCollection trCategoryCollection(m_trUser, "Category");
	return trCategoryCollection.Count();
}

int	TemplateSettingsHelper::GetTemplateCount(int nCategory)
{	
	TreeNode trCategory = getCategoryNode(nCategory);
	if( trCategory && trCategory.FirstNode )
		return trCategory.Children.Count();
	
	return 0;
}

bool TemplateSettingsHelper::HasUntitledCategory()
{	
	TreeNode trCategory = getCategoryNode(NUM_UNTITLED_CATEGORY_ID);
	if( trCategory )
	{
		return trCategory.DataID == NUM_UNTITLED_CATEGORY_ID;
	}
	
	return false;
}

void TemplateSettingsHelper::RemoveTemplate(int nCategory, int nTemplate, bool bDeleteFile, bool bResetNodeIDs, bool* pbUntitledCategoryRemoved)
{
	/// Iris 6/11/2015 ORG-13203-P3 FIX_FAIL_TO_SHOW_TEMPLATE_INFO_AFTER_CLICKED_SCAN_TEMPLATE
	bool bUntitledCategoryRemoved = false;
	///End FIX_FAIL_TO_SHOW_TEMPLATE_INFO_AFTER_CLICKED_SCAN_TEMPLATE
	
	TreeNode trCategory = getCategoryNode(nCategory);
	if( trCategory )
	{
		TreeNode trTemplate = getTemplateNode(trCategory, nTemplate);
		if( trTemplate )
		{			
			if( bDeleteFile )
			{
				string strFileName;
				
				//delete the template file only in the UFF
				if(  trTemplate.GetAttribute(STR_FILENAME_ATTRIB, strFileName) && strFileName.IsFile() 
					&& 0 == (GetFilePath(strFileName)).CompareNoCase(STR_ORIGIN_USER_FILE_FOLDER) )
				{
					string strDeleteFolder = STR_ORIGIN_USER_FILE_FOLDER + "DeletedTemplates\\";
					CheckMakePath(strDeleteFolder);
					
					string strDeletePath = strDeleteFolder + GetFileName(strFileName, false);

					if( strDeletePath.IsFile() )
						DeleteFile(strDeletePath);	
					
					MoveFile(strFileName, strDeletePath);
					
					// EMF file
					vector<string> vsImageExt = {".EMF", ".BMP"};
					for(int ii = 0; ii < vsImageExt.GetSize(); ii++)
					{
						string strImageFile = GetFilePath(strFileName) + GetFileName(strFileName, true) + vsImageExt[ii];						
						if( strImageFile.IsFile() )
						{
							string strDeleteImagePath = strDeleteFolder + GetFileName(strFileName, true) + vsImageExt[ii];
							
							if( strDeleteImagePath.IsFile() )
								DeleteFile(strDeleteImagePath);					
							MoveFile(strImageFile, strDeleteImagePath);
						}
					}					
				}
			}			
			
			trTemplate.Remove();
			
			/// Iris 6/3/2015 ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
			if( NUM_UNTITLED_CATEGORY_ID == trCategory.DataID && !trCategory.FirstNode )
			{
				trCategory.Remove();
				bUntitledCategoryRemoved = true; /// Iris 6/11/2015 ORG-13203-P3 FIX_FAIL_TO_SHOW_TEMPLATE_INFO_AFTER_CLICKED_SCAN_TEMPLATE
			}
			///End ADD_IMAGE_HINT_FOR_NO_TEMPLATE_MATCH_FILTER
			
			if( bResetNodeIDs )
				setupTemplateNodeID(trCategory);
			
			SetModifyFlag();
		}
	}
	
	/// Iris 6/11/2015 ORG-13203-P3 FIX_FAIL_TO_SHOW_TEMPLATE_INFO_AFTER_CLICKED_SCAN_TEMPLATE
	if( pbUntitledCategoryRemoved )
		*pbUntitledCategoryRemoved = bUntitledCategoryRemoved;
	///End FIX_FAIL_TO_SHOW_TEMPLATE_INFO_AFTER_CLICKED_SCAN_TEMPLATE
}


void TemplateSettingsHelper::InsertTemplate(int nCategory, int nTemplate, LPCSTR lpcszFilePath, bool bResetNodeIDs)
{
	TreeNode trCategory = getCategoryNode(nCategory);
	if( trCategory )
	{
		string str = "Template1";
		int index = tree_get_next_enum_tag_name(trCategory, str);	
		
		TreeNode trNewTemplate;
		TreeNode trTemplateBefore = getTemplateNode(trCategory, nTemplate);
		if( nTemplate < 0 || !trTemplateBefore )		
			trNewTemplate = trCategory.AddNode(str);
		else		
			trNewTemplate = trCategory.InsertNode(trTemplateBefore, str);			
	
		trNewTemplate.SetAttribute(STR_FILENAME_ATTRIB, lpcszFilePath);	
			
		trNewTemplate.SetAttribute(STR_CLONEABLE_ATTRIB, _is_template_cloneable(lpcszFilePath));
		
		trNewTemplate.SetAttribute(STR_LABEL_ATTRIB, GetFileName(lpcszFilePath, true));
		
		if( bResetNodeIDs )
			setupTemplateNodeID(trCategory);
		
		SetModifyFlag();
	}
}

int	TemplateSettingsHelper::AddTemplates(int nCategory, const vector<string>& vsFilePaths)
{
	TreeNode trCategory = getCategoryNode(nCategory);
	if(!trCategory)
	{
		trCategory = checkGetUntitledCategory();
		nCategory = NUM_UNTITLED_CATEGORY_ID;
	}
	
	int count;
	for(int index = 0; index < vsFilePaths.GetSize(); index++)
	{
		string strFilePath = vsFilePaths[index];
		///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
		if(strFilePath.IsEmpty())
			continue;
		///------ End IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
		if( !m_trUser.FindNodeByAttribute(STR_FILENAME_ATTRIB, strFilePath) )
		{
			if(!CheckSetTemplateName(strFilePath, nCategory, &count))	///------ Tony 08/12/2015 ORG-13290-S8 ADD_SCAN_TEMPLATE_DETECT_DUPLICATE_NAME
			{
				InsertTemplate(nCategory, -1, strFilePath, false);
				count++;
			}
		}
	}
	
	if( count > 0)
		setupTemplateNodeID(trCategory);
	
	return count;
}

bool TemplateSettingsHelper::GetTemplateInfo(int nCategory, int nTemplate, string& strDescription, string& strLocation, string& strDateModified)
{		
	GetTemplateFilePath(nCategory, nTemplate, strLocation);		
	
	if( strLocation.IsFile() )
	{		
		okutil_get_Origin_file_comment(strLocation, &strDescription); 		
	
		strDateModified = GetFileModificationDate(strLocation, LDF_SHORT_AND_HHMMSS_SEPARCOLON);
		
		return true;
	}
	
	return false;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////  TemplateLibraryDlg  ///////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
static int s_nFilterState = -1;	/// Nolan 12/03/2015 ORG-14036-S1 SYS_VAR_TO_SET_TEMPLATE_LIBRARY_FILTER_STATE

BOOL TemplateLibraryDlg::OnInitDialog()
{
	ResizeDialog::OnInitDialog(IDC_PREVIEW_LIST, STR_DLG_NAME);
	
	m_XMLHelper.Init(EXIST_GRAPH);		
	
	m_PreviewList.Init(IDC_PREVIEW_LIST, *this, &m_XMLHelper, m_strDefaultTemplate);			
	
	vector<uint> vnBtnIDs = {
		IDC_BUTTON_SCAN_USER_TEMPLATES
		, IDC_BUTTON_NEW_CATEGORY
		, IDC_BUTTON_ADD_TEMPLATE
		, IDC_BUTTON_FILTER	
	};
	vector<uint> vnBitmaps = {
		IDB_TEMPLATE_SCAN
		, IDB_CATEGORY_NEW
		, IDB_TEMPLATE_ADD
		, IDB_TEMPLATE_FILTER
	};
	vector<string> vsTips(vnBtnIDs.GetSize());
	vsTips[0] = _L("Scan User Template");
	vsTips[1] = _L("New Category");
	vsTips[2] = _L("Add Template");
	vsTips[3] = _L("Template Filter");	
	
	for(int ii = 0; ii < vnBtnIDs.GetSize(); ii++)
	{
		BitmapRadioButton btn = GetItem(vnBtnIDs[ii]);
		vector<string> vs(1);
		vs[0] = vsTips[ii];
		
		bool bIsFilterButton = ( IDC_BUTTON_FILTER == vnBtnIDs[ii] );
		UINT nStates = bIsFilterButton ? 2 : 1;
		btn.Init(nStates, vnBitmaps[ii], 16, vs, 0);
		
		if( bIsFilterButton )
		/// Nolan 12/03/2015 ORG-14036-S1 SYS_VAR_TO_SET_TEMPLATE_LIBRARY_FILTER_STATE
		//	btn.Check = LoadSetting("FilterOn", 0, STR_DLG_NAME);
		{
			double dTLF = -1;
			LT_get_var("@TLF", &dTLF);
			
			if (dTLF == 0)
				s_nFilterState = false;
			else if (dTLF == 2)
			{
				s_nFilterState = LoadSetting("FilterOn", 0, STR_DLG_NAME);
			}
			else if (s_nFilterState == -1)
			{
				s_nFilterState = false;
			}
			
			btn.Check = s_nFilterState;
		}
		/// End SYS_VAR_TO_SET_TEMPLATE_LIBRARY_FILTER_STATE
	}
	m_btnFilter = GetItem(IDC_BUTTON_FILTER); 
	
	setFilterControlsEnable();
	m_PreviewList.SetFilterOption(m_btnFilter && m_btnFilter.Check);	
	
	initErrorMessageControl(); ///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	
	SetInitReady();	
	
	return TRUE;
}

BOOL TemplateLibraryDlg::OnReady()
{
	m_PreviewList.SetupScrollBarPos(s_hWndPreviewList, 0);
	
	return TRUE;
}

/// Iris 5/27/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTONS
void TemplateLibraryDlg::updatePlottingButtonEnableStatus()
{		
	Button btnPlotSetup = GetItem(IDC_BUTTON_PLOT_SETUP);
	///------ Tony 09/16/2015 ORG-12495-S19 DONT_CLOSE_TL_AFTER_CLICK_PLOT
	//Button btnPlot = GetItem(IDOK);
	Button btnPlot = GetItem(IDC_BUTTON_PLOT);
	///------ End DONT_CLOSE_TL_AFTER_CLICK_PLOT
	Button btnOriginalData = GetItem(IDC_BTN_ORIGINAL_DATA);	///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	
	string strTemplate;
	int nRow, nCol;
	if( m_PreviewList.GetSelCell(&nRow, &nCol) )		
		strTemplate = m_PreviewList.GetTemplateFilePath(nRow, nCol);
	
	if( strTemplate.IsEmpty() || !strTemplate.IsFile() )
	{		
		btnOriginalData.Enable = 	///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
		btnPlotSetup.Enable = 
		btnPlot.Enable = false;
	}
	else
	{
		btnPlotSetup.Enable = true;
		///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
		//btnPlot.Enable = has_data_range_selected() || m_btnFilter && m_btnFilter.Check || m_PreviewList.CheckByFilter(strTemplate);
		bool bCloneable = _is_template_cloneable(strTemplate);
		
		btnPlot.Enable = !bCloneable && has_data_range_selected() || bCloneable && m_PreviewList.CheckByFilter(strTemplate);
	
		btnOriginalData.Enable = bCloneable;		
		///------ End IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	}
}
///End IMPROVE_PLOT_BUTTONS

///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
void TemplateLibraryDlg::checkErrorMessage(int nCheckErrCntrl)
{
	if( !IsInitReady() )
		return;
	
	string strErrMsg;
	if( CHECK_ERROR_CNTRL_TEMPLATE_NOT_MATCH & nCheckErrCntrl)
	{
		int nErrMsg = CER_DO_NOT_MATCH_TEMPLATE;
		
		//remove the old one and set the new one on the end of the vector which show firstly
		vector<uint> vn;
		m_vnErrMsgs.Find(MATREPL_TEST_EQUAL, nErrMsg, vn); // to check if the current error already existed in error array.	
		for(int ii = vn.GetSize() - 1; ii >= 0; ii--)
		{
			m_vnErrMsgs.RemoveAt(vn[ii]);
		}
		
		string strTemplate;
		int nRow, nCol;
		if( m_PreviewList.GetSelCell(&nRow, &nCol) )		
			strTemplate = m_PreviewList.GetTemplateFilePath(nRow, nCol);
		
		if( !strTemplate.IsEmpty() )
		{
			if( !_is_template_match_with_active_sheet(strTemplate) && _is_template_cloneable(strTemplate) )
				m_vnErrMsgs.Add(nErrMsg);
		}
	}
	
	//show the last msg
	if(m_vnErrMsgs.GetSize() > 0)
	{
		string strErrMsg;
		ocu_load_err_msg_str(m_vnErrMsgs[m_vnErrMsgs.GetSize() - 1], &strErrMsg);

		setErrorMessage(strErrMsg);
	}
	else
	{
		setErrorMessage(NULL);
	}
	
}

void TemplateLibraryDlg::resizeDlg()
{
	RECT rDlg;
	Window wnd = GetWindow();
	wnd.GetClientRect(&rDlg);
	int cx = RECT_WIDTH(rDlg);
	int cy = RECT_HEIGHT(rDlg);
	OnDlgResize(0, cx, cy);	
}

bool TemplateLibraryDlg::setErrorMessage(LPCSTR lpcstrMsg)
{
	Control ctrlErrMsg = GetItem(IDC_ERR_MESSAGE_BOX);

	bool bOldVisible = ctrlErrMsg.Visible;
	
	if(lpcstrMsg || lpcstrMsg == "")
	{
		string strCurErrMsg = ctrlErrMsg.Text;
		if(0 != strCurErrMsg.Compare(lpcstrMsg))
			ctrlErrMsg.Text = lpcstrMsg;
		
		if(!bOldVisible)
			ctrlErrMsg.Visible = true;
	}
	else
	{
		ctrlErrMsg.Text = "";
		
		if(bOldVisible)
			ctrlErrMsg.Visible = false;
	}

	if(bOldVisible != ctrlErrMsg.Visible)
		resizeDlg();

	return true;
}

void TemplateLibraryDlg::initErrorMessageControl()
{
	RichEdit ctrlErrMsg = GetItem(IDC_ERR_MESSAGE_BOX);
	ctrlErrMsg.Text = "";
	ctrlErrMsg.Visible = false;
	
	CHARFORMAT cf; 	
	ctrlErrMsg.GetDefaultCharFormat(cf);	
	O_ADD_BIT(cf.dwEffects, CFE_BOLD);
	O_ADD_BIT(cf.dwMask, CFM_BOLD);
	
	O_REMOVE_BIT(cf.dwEffects, CFE_AUTOCOLOR);
	O_ADD_BIT(cf.dwMask, CFM_COLOR);
	cf.crTextColor = RGB(255, 0, 0);
	ctrlErrMsg.SetDefaultCharFormat(cf);
}
///------ End IMPROVE_PLOT_BUTTON_WITH_CLONEABLE

BOOL TemplateLibraryDlg::OnDlgResize(int nType, int cx, int cy)
{
	if( !IsInitReady() )
		return FALSE;
	
	MoveControlsHelper	_temp(this);
	int nGap = GetControlGap();	
	
	// move bottom buttons
	RECT rrOK, rrCancel;
	///------ Tony 09/16/2015 ORG-12495-S19 DONT_CLOSE_TL_AFTER_CLICK_PLOT
	//GetControlClientRect(IDOK, rrOK);	
	GetControlClientRect(IDC_BUTTON_PLOT, rrOK);	
	///------ End DONT_CLOSE_TL_AFTER_CLICK_PLOT
	
	uint nBottomRightBtnIDs[] = {
		IDCANCEL		
		///------ Tony 09/16/2015 ORG-12495-S19 DONT_CLOSE_TL_AFTER_CLICK_PLOT
		//, IDOK
		, IDC_BUTTON_PLOT
		///------ End DONT_CLOSE_TL_AFTER_CLICK_PLOT
		, IDC_BUTTON_PLOT_SETUP		
		, IDC_BTN_ORIGINAL_DATA
		, 0 };
	ArrangeControlsRightLeft(nBottomRightBtnIDs, cx, cy - nGap - RECT_HEIGHT(rrOK));
		
	///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	Control ctrlErrMsg;
	RECT rrErrMsg;
	GetControlClientRect(IDC_ERR_MESSAGE_BOX, rrErrMsg, &ctrlErrMsg);	
	bool bShowErrMsg = ctrlErrMsg && ctrlErrMsg.Visible;
	double dNextHeight = RECT_HEIGHT(rrOK) + 2 * nGap;
	if( bShowErrMsg )
	{		
		rrErrMsg.left = nGap;
		rrErrMsg.right = cx - nGap;
		int nWidth = RECT_WIDTH(rrErrMsg);				
		int nHeight = ctrlErrMsg.Measure(ctrlErrMsg.Text, &nWidth, false);
		rrErrMsg.bottom = cy - dNextHeight;
		rrErrMsg.top = rrErrMsg.bottom - nHeight;
		MoveControl(ctrlErrMsg, rrErrMsg);
		
		dNextHeight += nHeight + nGap;
	}
	///------ End IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	
	// move static and edit boxes
	uint nIDs[] = {
		IDC_STATIC_DESCRIPTION
		, IDC_EDIT_DESCRIPTION
		, IDC_STATIC_LOCATION
		, IDC_EDIT_LOCATION
		, IDC_STATIC_DATE_MODIFIED
		, IDC_EDIT_DATE_MODIFIED
		, 0 };
	RECT rrEdit;
	GetControlClientRect(IDC_EDIT_DATE_MODIFIED, rrEdit);
	///------ Tony 07/02/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	//int dy = cy - (rrEdit.bottom + RECT_HEIGHT(rrOK) + 2 * nGap);	
	int dy = cy - (rrEdit.bottom + dNextHeight);	
	///------ End IMPROVE_PLOT_BUTTON_WITH_CLONEABLE
	MoveControls(nIDs, 0, dy);	
	GetControlClientRect(IDC_EDIT_DESCRIPTION, rrEdit); // get the rectange of editbox after moving
	
	uint nEditIDs[] = {
		IDC_EDIT_DESCRIPTION
		, IDC_EDIT_LOCATION
		, IDC_EDIT_DATE_MODIFIED
		, 0};
	int index = 0;
	while( nEditIDs[index] > 0 )
	{
		RECT rr;
		Control ed;
		GetControlClientRect(nEditIDs[index], rr, &ed);
		rr.right = cx - nGap;
		MoveControl(ed, rr);
		
		index++;
	}
	
	
	RECT rrStaticTemplate;
	Control ctrlStaticTemplate;
	GetControlClientRect(IDC_STATIC_TEMPLATE, rrStaticTemplate, &ctrlStaticTemplate);	

	int nHintWidth = RECT_WIDTH(rrStaticTemplate);
	int nHintHeidht = RECT_HEIGHT(rrStaticTemplate);
	rrStaticTemplate.top = nGap;
	rrStaticTemplate.bottom = rrStaticTemplate.top + nHintHeidht;
	rrStaticTemplate.left = nGap;
	rrStaticTemplate.right = rrStaticTemplate.left + nHintWidth;
	
	MoveControl(ctrlStaticTemplate, rrStaticTemplate);
	
	// move the bitmap radio buttons
	uint nBtnIDs[] = {
		IDC_BUTTON_SCAN_USER_TEMPLATES
		, IDC_BUTTON_NEW_CATEGORY
		, IDC_BUTTON_ADD_TEMPLATE
		, IDC_BUTTON_FILTER	
		, 0 };
	
	int nRight = rrStaticTemplate.right;
	int nBtnWidth = 25, nBtnHeight = 25;
	int nBtnGap = 5;	
	nBtnWidth = check_convert_size_with_DPI(GetWindow(), nBtnWidth, true);
	nBtnHeight = check_convert_size_with_DPI(GetWindow(), nBtnHeight, false);
	nBtnGap = check_convert_size_with_DPI(GetWindow(), nBtnGap, true);
	
	int ii = 0;
	while ( nBtnIDs[ii] > 0 )
	{
		RECT rrBtn;		
		rrBtn.top = nGap;
		rrBtn.bottom = rrBtn.top + nBtnHeight;		
		rrBtn.left = nRight + nBtnGap;		
		rrBtn.right = rrBtn.left + nBtnWidth;
		
		Button btn = GetItem( nBtnIDs[ii] );
		MoveControl(btn, rrBtn);
		
		nRight = rrBtn.right;		
		ii++;
	}
	
	RECT rrPreview;
	Control ctrlPreview;
	GetControlClientRect(IDC_PREVIEW_LIST, rrPreview, &ctrlPreview);
	rrPreview.left = nGap;
	rrPreview.right = cx - nGap;	
	rrPreview.top = rrStaticTemplate.bottom + nGap / 2;
	rrPreview.bottom = rrEdit.top - nGap;	
	
	MoveControl(ctrlPreview, rrPreview);
	
	// after dlg resize ready, icon list and preview list can start to load icon/image according to the width of list. 
	// The width of list decided the number of columns in the two lists.	
	m_PreviewList.OnDlgResize(RECT_WIDTH(rrPreview), RECT_HEIGHT(rrPreview));
	
	updateControlWithTemplateInfo();
	
	return TRUE;
}

BOOL TemplateLibraryDlg::OnRestoreSize(ODWP dwSizeInfo)
{
	void * p = (void*)dwSizeInfo;
	DLGSIZEINFO *pSz = (DLGSIZEINFO*)p;
	lstrcpyn(pSz->szDialogName, STR_DLG_NAME, MAXLINE);

	SIZE sz;
	GetDlgOptimalSize(sz);
	
	pSz->top = -1;
	pSz->left = -1;
	pSz->width = 760;
	pSz->height = 600;
	
	return TRUE;
}

void TemplateLibraryDlg::OnPreviewListAfterEdit(Control flxControl, int nRow, int nCol)
{
	if( m_PreviewList.IsCategoryRow(nRow) )
	{
		int nCategory = m_PreviewList.GetCategoryIndex(nRow);
		
		string str = m_PreviewList.GetCell(nRow, nCol);
		string strOld = str;
		m_XMLHelper.RenameCategory(nCategory, str);

		//reset cell because m_XMLHelper.RenameCategory may change str
		if( 0 != strOld.CompareNoCase(str) )
		{
			for(int nn = 0; nn < m_PreviewList.GetNumCols(); nn++)
			{
				m_PreviewList.SetCell(nRow, nn, str);
			}
		}
	}
}

void TemplateLibraryDlg::OnPreviewListCellSelChange(Control ctrl)
{
	///------ Tony 07/13/2015 ORG-13290-S3 IMPROVE_HINT_IMAGE
	if( 0 == m_PreviewList.GetNumRows() || m_PreviewList.IsTemplateNone()) 
		return;
	///------ End IMPROVE_HINT_IMAGE
	
	updateControlWithTemplateInfo();
	

	checkErrorMessage(CHECK_ERROR_CNTRL_TEMPLATE_NOT_MATCH);
	
	updatePlottingButtonEnableStatus(); /// Iris 5/27/2015 ORG-12495-S10 IMPROVE_PLOT_BUTTONS	
}

void TemplateLibraryDlg::OnListBeforeEdit( Control ctrl, int nRow, int nCol, BOOL* pCancel )
{		
	bool bIsCategoryRow = m_PreviewList.IsCategoryRow(nRow);
	
	if( !bIsCategoryRow || 0 != nCol && bIsCategoryRow )
		*pCancel = true;
}


void	TemplateLibraryDlg::OnListToolBarClick(Control ctrl, int nRow, int nCol, int nMouseOverIndex)
{
	if ( GetDlgCtrlID(ctrl.GetSafeHwnd()) == IDC_PREVIEW_LIST )
	{
		switch(nMouseOverIndex)
		{
		case TOOL_BAR_INDEX_EDIT:
			m_PreviewList.DoEdit(nRow, nCol);
			break;
			
		case TOOL_BAR_INDEX_DELETE:
			//delete the template in specified cell		
			m_PreviewList.Delete(nRow, nCol);
			break;
			
		default:
			ASSERT(false);
			break;
		}
		
		//after delete or edit, should reset status by sel change
		OnPreviewListCellSelChange(ctrl);///------ Tony 07/23/2015 ORG-13290-S5 INITIALIZE_THE_CATEGORY_NAME_IMPROVE
	}
}

void TemplateLibraryDlg::OnPreviewListStartDrag(Control ctrl, int* pnAllowEffect, int* pnRow, int* pnCol, int* pnRowSel, int* pnColSel)
{	
	if( NULL != pnAllowEffect && NULL != *pnRow && NULL != pnCol )
		*pnAllowEffect = m_PreviewList.HasImage(*pnRow, *pnCol) ? flexDropEffectMove : flexDropEffectNone;	
}

void TemplateLibraryDlg::OnPreviewListDragFeedback(Control ctrl, int* pnAllowEffect, bool* pCancel)
{	
	bool bAllow = m_PreviewList.IsAllowDrop();
		
	if( pnAllowEffect )
		*pnAllowEffect = bAllow ? flexDropEffectMove : flexDropEffectNone;
}

void TemplateLibraryDlg::OnPreviewListDragDropFinish(Control ctrl, int* pnAllowEffect)
{	
	m_PreviewList.DragAndDrop();	
}

void TemplateLibraryDlg::OnBtnScanUserTemplatesClick(Control ctrl)
{
	bool bGot = m_XMLHelper.ScanUserDefinedTemplate();
	
	if( bGot )
	{
		m_PreviewList.Load();
	
		updateControlWithTemplateInfo();
	}
}

void TemplateLibraryDlg::OnClickNewCategory(Control ctrl)
{
	string strCategoryLabel;
	m_XMLHelper.NewCategory(strCategoryLabel);
	
	m_PreviewList.NewCategory(strCategoryLabel);
	m_PreviewList.SetupScrollBarPos(s_hWndPreviewList, m_PreviewList.GetNumRows());
	
	if( m_PreviewList.IsTemplateNone() && !m_PreviewList.IsFilterOn() )
		m_PreviewList.Load();
}

void TemplateLibraryDlg::OnBtnAddTemplateClick(Control ctrl)
{
	s_bOnActivateReady = false;
	
	vector<string> vsFilePaths;
	string		strExts =  m_XMLHelper.GetTemplateExt();
	StringArray	saExts;
	int			nn = strExts.GetTokens(saExts, '|');
	string		strFileDlgFilters;
	for (int ii = 0; ii < nn; ii++)
	{
		if (0 < ii)
			strFileDlgFilters += ';';
		string		strOneFilter;
		strOneFilter.Format("*.%s", saExts[ii]);
		strFileDlgFilters += strOneFilter;
	}
	GetMultiOpenBox(vsFilePaths, strFileDlgFilters, STR_ORIGIN_USER_FILE_FOLDER);
	
	if( vsFilePaths.GetSize() > 0 && m_XMLHelper.AddTemplates(m_PreviewList.GetCurrentCategory(), vsFilePaths) > 0 )
		m_PreviewList.Load();
	
	s_bOnActivateReady = true;
}

static void _check_open_image_profile_dlg(GraphPage& gp)
{
	GraphLayer gl = gp.Layers(0);
	if (gl)
	{
		if (gl.GetBinaryStorage(STR_IMAGE_PROFILE_STORAGE, NULL))
		{
			GraphObject btnProfile = gl.GraphObjects("ProfileButton");
			if (btnProfile)
			{
				string strScript;
				if (get_go_LTScript_event(btnProfile, strScript))
				{
					LT_execute(";" + strScript);
				}
			}
		}
	}
}

BOOL TemplateLibraryDlg::OnDblClickToPlot(Control ctrl)
{
	Button btnPlot = GetItem(IDC_BUTTON_PLOT);
	if (btnPlot.Enable)
		OnClickPlotButton(btnPlot);
	return TRUE;
}


BOOL TemplateLibraryDlg::OnClickPlotButton(Control ctrl)
{
	string strTemplate = getCurrentSelectedTemplate(false);

	if( strTemplate.IsEmpty() )
		return false;	

	/// Iris 4/10/2015 ORG-12495-S7 ONLY_SHOW_USER_DEFINED_TEMPLATE		
	Page pg = Project.Pages();
	if( pg )
	{
		///Kyle 07/07/2015 ORG-13291-P1 PLOT_FROM_USER_TEMPLATE_SUPPORT_EXCEL
		//int nType = pg.GetType();
		//if( EXIST_WKS != nType && EXIST_MATRIX != nType ) // if the active window is NOT wks or mat, need to open Plot Setup dialog
		if( !_is_page_support_auto_plotting(pg.GetType()) )
		///End PLOT_FROM_USER_TEMPLATE_SUPPORT_EXCEL
		{
			Control ctrl = GetItem(IDC_BUTTON_PLOT_SETUP);
			return OnClickPlotSetupButton(ctrl);
		}
	}	
	///End ONLY_SHOW_USER_DEFINED_TEMPLATE		
	
	if( !doSmartPlot() )
	{	
		string strCmd;
		make_plot_command_with_template(strTemplate, IDM_PLOT_UNKNOWN, strCmd);
		
		if (LT_execute(strCmd))
		{
			GraphPage gp = Project.Pages();
			if (gp)
			{
				_check_open_image_profile_dlg(gp);
			}
		}
	}
	
	GraphPage gpNew = Project.Pages();
	if( pg && gpNew && pg.GetName() != gpNew.GetName() )
	{
		gpNew.InvokePresetNames();
	}
	
	return TRUE;
}


bool TemplateLibraryDlg::doSmartPlot()
{
	string strTemplate = getCurrentSelectedTemplate(false);

	int nMatchBy;
	if( m_PreviewList.CheckByFilter(strTemplate) )
	{		
		string strCmd;
		strCmd.Format("worksheet -pa ? \"%s\"", strTemplate);
		if( LT_execute(strCmd) )
		{			
			//SendMessage(WM_CLOSE);	///------ Tony 09/16/2015 ORG-12495-S19 DONT_CLOSE_TL_AFTER_CLICK_PLOT
			return true;
		}
	}
	
	return false;
}


BOOL TemplateLibraryDlg::OnClickPlotSetupButton(Control ctrl)
{	
	FUNC_PLOT_SETUP pfn = Project.FindFunction( "PlotSetup", "OriginLab\\PlotSetup.c" );
	ASSERT( pfn );
	if( !pfn )
		return FALSE;	

	HWND hWndParent = GetSafeHwnd();
	ASSERT( hWndParent );
	
	if( pfn( true, 0, getCurrentSelectedTemplate(false), hWndParent ) )
	{
		//PostMessage(WM_CLOSE);	///------ Tony 09/16/2015 ORG-12495-S19 DONT_CLOSE_TL_AFTER_CLICK_PLOT
		return true;
	}
	return false;
}

void TemplateLibraryDlg::OnActivate(UINT nState, HWND hwndOther, BOOL bMinimized)
{
	if(!IsInitReady() ||!s_bOnActivateReady)
	
		return;
	s_bOnActivateReady = false;
	bool bRooledup = IsRolledup();
	if((WA_ACTIVE == nState || WA_CLICKACTIVE == nState) && !bMinimized && !IsRolledup())
	{
		m_PreviewList.OnActivate()
	}
	else if(WA_INACTIVE == nState)
	{
		m_PreviewList.Save();
	}
	
	s_bOnActivateReady = true;

}

BOOL TemplateLibraryDlg::OnChangePage()
{
	if(!IsInitReady())
		return FALSE;

	Page pg = Project.Pages();
	if( !pg || !_is_page_support_auto_plotting(pg.GetType()) )
	{
		Visible = FALSE;
	}
	else
	{
		Visible = TRUE;
	}
	return TRUE;
}

bool TemplateLibraryDlg::OnSystemCommand(int nCmd)
{
	if(SC_MINIMIZE == nCmd && IsRolledup())
		m_PreviewList.Load();
	
	return ResizeDialog::OnSystemCommand(nCmd);
}

BOOL TemplateLibraryDlg::OnActiveLayerChange()
{
	checkErrorMessage(CHECK_ERROR_CNTRL_TEMPLATE_NOT_MATCH);
	updatePlottingButtonEnableStatus();
	
	return TRUE;
}


BOOL TemplateLibraryDlg::OnSelectionChange()
{
	checkErrorMessage(CHECK_ERROR_CNTRL_TEMPLATE_NOT_MATCH);
	updatePlottingButtonEnableStatus();
	return TRUE;
}

BOOL TemplateLibraryDlg::OnTemplateUpdate(WPARAM wParam, LPARAM lParam)
{
	string strTemplate = s_strUpdatedTemplateFile;
	if ( strTemplate.IsFile() )
	{
		s_strUpdatedTemplateFile.Empty(); //clear
		m_PreviewList.OnTemplateUpdate(strTemplate);
	}
	return TRUE;
}

string TemplateLibraryDlg::getCurrentSelectedTemplate(bool bOnlyName)
{
	int nRow, nCol;
	int nCategory, nTemplate;

	m_PreviewList.GetSelCell(&nRow, &nCol);
	nCategory = m_PreviewList.GetCategoryIndex(nRow);
	nTemplate = m_PreviewList.GetTemplateIndex(nRow, nCol);
	
	string strFile;
	m_XMLHelper.GetTemplateFilePath(nCategory, nTemplate, strFile);
	if( strFile.IsFile() && bOnlyName )
	{
		string strName = GetFileName(strFile, true);
		return strName;
	}
	return strFile;
}

BOOL TemplateLibraryDlg::OnDestroy()
{
	m_XMLHelper.Save();
	
	double dTLF = -1;
	LT_get_var("@TLF", &dTLF);
	s_nFilterState = m_btnFilter.Check;
	
	if (dTLF == 2)
		SaveSetting("FilterOn", m_btnFilter.Check, STR_DLG_NAME);

	delete this;
	s_pTemplateLibraryDlg = NULL;

	return TRUE;
}

BOOL TemplateLibraryDlg::OnFilterClicked(Control ctrl)
{	
	updatePreviewListByFilterOption(true);
	return TRUE;
}

BOOL TemplateLibraryDlg::OnBtnGetOriginalData(Control ctrl)
{
	/// Iris 6/30/2015 ORG-12495-S16 ADD_GET_ORIGINAL_DATA_BUTTON
	string strTemplate = getCurrentSelectedTemplate(false);
	ASSERT( strTemplate.IsFile() );
	
	string strWinName;
	Project.CreateSourceBook(strTemplate, strWinName);
	
	Page pg(strWinName);
	if( pg && !pg.Show )
		pg.Show = TRUE;
	///End ADD_GET_ORIGINAL_DATA_BUTTON
	
	return TRUE;
}

void TemplateLibraryDlg::OnGridMouseMove(Control ctrl, short nButton, short nShift, float fX, float fY)
{
	int	nRow, nCol;
	if ( !m_PreviewList.GetMouseCell(nRow, nCol) || nRow < 0 || nCol < 0 )
		return;
	string strTips = m_PreviewList.GetCell(nRow, nCol);
	PictureHolder pict;
	if ( !m_PreviewList.GetCellPicture(nRow, nCol, pict) || pict.GetHeight() <= 0 )
		strTips = ""; //only display for those contains picture.
	m_PreviewList.SetToolTipsText(strTips);
}


void TemplateLibraryDlg::setFilterControlsEnable()
{
	bool bEnableFilter = false;
	Page pg = Project.Pages();
	if( pg )
	{
		///Kyle 07/07/2015 ORG-13291-P1 PLOT_FROM_USER_TEMPLATE_SUPPORT_EXCEL
		//int nType = pg.GetType();
		//bEnableFilter = (EXIST_WKS == nType || EXIST_MATRIX == nType);
		bEnableFilter = _is_page_support_auto_plotting(pg.GetType());
		///End PLOT_FROM_USER_TEMPLATE_SUPPORT_EXCEL
	}
	
	Button btnFilter = GetItem(IDC_BUTTON_FILTER);	 
	if( btnFilter )	
		btnFilter.Enable = bEnableFilter;			
}

void TemplateLibraryDlg::updatePreviewListByFilterOption(bool bUpdateList)
{		
	bool bUpdated = m_PreviewList.SetFilterOption(m_btnFilter && m_btnFilter.Check);	
	
	if( bUpdateList && bUpdated )
		m_PreviewList.Load();
}

void TemplateLibraryDlg::updateControlWithTemplateInfo()
{	
	bool bIsCategoryTitleRow;
	bool bHasImage;
	
	int nCategory, nTemplate;
	
	if( 0 == m_PreviewList.GetNumRows() )
		return;
	
	int nx, ny, row, col;
	m_PreviewList.GetSelCell(nx, ny, row, col);		
	if( row < 0 || col < 0 )
		return;
	
	bIsCategoryTitleRow = m_PreviewList.IsCategoryRow(row);
	bHasImage = m_PreviewList.HasImage(row, col);
	
	if( !bIsCategoryTitleRow )
	{
		nCategory = m_PreviewList.GetCategoryIndex(row);
		nTemplate = m_PreviewList.GetTemplateIndex(row, col);
	}	

	string strDescription, strLocation, strDateModified;
	if( !bIsCategoryTitleRow && bHasImage && nTemplate >= 0 )
		m_XMLHelper.GetTemplateInfo(nCategory, nTemplate, strDescription, strLocation, strDateModified);
	
	Edit edDescription = GetItem(IDC_EDIT_DESCRIPTION);
	Edit edLocation = GetItem(IDC_EDIT_LOCATION);
	Edit edDateModified = GetItem(IDC_EDIT_DATE_MODIFIED);
	
	edDescription.Text = strDescription;
	edLocation.Text = strLocation;
	edDateModified.Text = strDateModified;
}

static bool _is_page_support_auto_plotting(int nType)
{
	switch( nType )
	{
	case EXIST_WKS:
	case EXIST_EXTERN_WKS:
	case EXIST_MATRIX:
		return true;
	}

	return false;
}

void	OpenTemplateLibraryDialog(int nPlotCategoryIndex)
{
	BOOL bUserDefined = FALSE;
	string strSelCategory = "";
	if ( PLOTCAT_USERDEFINED == nPlotCategoryIndex )
		bUserDefined = TRUE;
	else
	{
		string strCat = STR_CATEGORYS;
		strSelCategory = strCat.GetToken(nPlotCategoryIndex, '|');
	}
	OpenTemplateLibraryDlg(TRUE, bUserDefined, strSelCategory);
	return;
}

void OpenTemplateLibraryDlg(BOOL bMakePlot = TRUE, BOOL bIsUserDefined = TRUE, string strCategory = "", string strTemplate = "", int nPlotID = IDM_PLOT_UNKNOWN)
{
	if(!s_pTemplateLibraryDlg)
	{
		s_pTemplateLibraryDlg = new TemplateLibraryDlg(bMakePlot, bIsUserDefined, strCategory, strTemplate, nPlotID);
		s_pTemplateLibraryDlg->Create(GetWindow());
	}
	else
	{
		SetActiveWindow(s_pTemplateLibraryDlg->GetSafeHwnd());
	}
}


void	NotifyTemplateLibraryOnSaveTemplate(string strTempateFullPath)
{
	if ( s_pTemplateLibraryDlg )
	{
		s_strUpdatedTemplateFile = strTempateFullPath;
		s_pTemplateLibraryDlg->PostMessage(WM_TEMPLATE_UPDATE, 0, 0);
	}

}


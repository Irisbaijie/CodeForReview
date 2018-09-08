#ifndef TEMPLATE_LIBRARY_DLG_H
#define TEMPLATE_LIBRARY_DLG_H


#define STR_DLG_NAME						"Template Library"

#define STR_UNTITLED_CATEGORY_NAME			"Category0"
#define NUM_UNTITLED_CATEGORY_ID			0

enum
{
	CATEGORY_USER_DEFINED = 0,
	
	CATEGORY_LINE,
	CATEGORY_SYMBOL,
	CATEGORY_LINE_AND_SYMBOL,	
	CATEGORY_COLUMN,
	
	CATEGORY_MULTI_Y,
	CATEGORY_Y_OFFSET,
	CATEGORY_MULTI_PANEL,
	
	CATEGORY_3D_XYY,
	CATEGORY_3D_SURFACE,
	CATEGORY_3D_SYMBOL,
	
	CATEGORY_STATS,
	CATEGORY_AREA,
	CATEGORY_CONTOUR,
	CATEGORY_PROFILE,	
	CATEGORY_SPECIALIZED,
	CATEGORY_STOCK,
	
	CATEGORY_ITEM_COUNT
};

#define SYS_CATEGORY_INDEX_OFFSET	CATEGORY_LINE

enum
{
	TYPE_SYSTEIM,
	TYPE_USER_DEFINED	
};

enum
{
	FILTER_BY_WIN_AUTO = 0,
	FILTER_BY_WIN_SHEET,	
	FILTER_BY_WIN_BOOK,
};

enum
{
	FILTER_BY_COL_INDEX = 0,
	FILTER_BY_COL_SN,
	FILTER_BY_COL_LN
};


enum
{
	CHECK_ERROR_CNTRL_TEMPLATE_NOT_MATCH = 0x0001;
};

class TemplateSettingsHelper
{
public:
	TemplateSettingsHelper(){}	
	~TemplateSettingsHelper() {}	
	
	bool	Init(int nPageType);
	bool 	NewCategory(string& strCategoryLabel);
	bool	RenameCategory(int index, string& strNewName);
	bool 	CheckCategory(LPCSTR lpcszTemplateFile, int* pnCategoryID = NULL, string* pstrCategoryName = NULL);	
	void 	ReorderCategory(int nCurrent, int nNew);
	bool	RemoveCategory(int index);
	
	bool 	GetCategoryLabel(int index, string& strLabel);
	
	bool	CheckSetTemplateName(LPCSTR lpcszTemplateFile, int nCategory, int* pnCount);
	
	bool 	GetTemplateFilePath(int nCategoryIndex, int nTemplateIndex, string& strFilename);	
	
	void	SetModifyFlag();
	bool	IsModified();
	bool	Save();
	bool	LoadXML();
	
	int		GetCategoryCount();	
	int		GetTemplateCount(int nCategory);		
	
	bool	HasUntitledCategory();

	void	RemoveTemplate(int nCategory, int nTemplate, bool bDeleteFile, bool bResetNodeIDs, bool* pbCategoryReomoved = NULL);
	
	void	InsertTemplate(int nCategory, int nTemplate, LPCSTR lpcszFilePath, bool bResetNodeIDs);
	int		AddTemplates(int nCategory, const vector<string>& vsFilePaths);
	
	bool 	ScanUserDefinedTemplate();
	
	bool 	GetTemplateInfo(int nCategory, int nTemplate, string& strDescription, string& strLocation, string& strDateModified);

	string 	GetTemplateExt() { return "otpu|otp"; }

	enum {
		AST_GET = 0,
		AST_SET,
		AST_DELETE,
	};
	void	AccessSuspendingTemplate(int nCategory, int nTempalteIndex, int nOpt = AST_DELETE, string& strTemplateFileName = NULL);
    
	
private:
	TreeNode getTree(bool bIsUser) { return bIsUser ? m_trUser : m_trSys; }
	
	TreeNode getCategoryNode(int index);
	TreeNode getTemplateNode(TreeNode& trCategory, int index);
	
	void	setupCategoryNodeID(TreeNode& tr);
	void 	setupTemplateNodeID(TreeNode& trCategory);


	TreeNode checkGetUntitledCategory();

private:
	Tree			m_trSys; // only read from OriginExe Folder\template.xml file, will not modify.
	Tree			m_trUser; // load from User File Folder\template.xml file, can be modify.
	Tree			m_trSuspending;
	int				m_nPageType;
	bool			m_bModified;
};

class TemplateListBase : public GridListControl
{
public:
	TemplateListBase() {}
	~TemplateListBase() {}
	
	virtual	void Init(int nID, WndContainer& dlg, TemplateSettingsHelper* pXMLHelper);
	
	virtual void AddTextRow(int nRow, LPCSTR lpcstrText, bool bHideMouseOverIcon = false);	

	virtual bool Load(bool bReloadXML = false) { ASSERT(FALSE); return false; }
	virtual bool Save() { ASSERT(FALSE); return false; }


	bool IsCategoryRow(int nRow);
	
	bool HasImage(int nRow, int nCol);
	
	void SetupScrollBarPos(HWND hWnd, int nRow);
	
	void OnDlgResize(int cx, int cy);
	
	virtual int GetCategoryIndex(int nRow);
	

	void	SetupMouseOverIcon(int nHighlight = 0);
	void	SetupCellIndicator(int nHighlight = 0);

	
protected:	
	int CalcColNum();
	
	void SetScrollBarPos(HWND hWnd, int nPos, bool bRefresh = true);
	
	double GetScrollBarPosPercent(HWND hWnd);	
	void SetScrollBarPosPercent(HWND hWnd, double dPosPercent, bool bRefresh = true);
	
	bool BeforeLoad(HWND hWndList);
	void AfterLoad(HWND hWndList, bool bReload);
	
	void SetCellHighlightEnable(bool bEnable);
	
	TemplateSettingsHelper*			m_pXMLHelper;
	
private:
	virtual int getCellWidth() { ASSERT(FALSE); return 0; }
	
private:
	bool				m_bIsReload;
	double				m_dScrollBarPosPercent;
};

class PreviewList : public TemplateListBase
{
public:
	PreviewList()
	{
	}
	~PreviewList()
	{
	}
	
	void Init(int nID, WndContainer& dlg, TemplateSettingsHelper* pXMLHelper, LPCSTR lpcszDefaultTemplate = NULL);
		
	bool NewCategory(LPCSTR strCategoryLabel);		
	
	bool IsAllowDrop();	
	void DragAndDrop();
	
	int GetCurrentCategory();
	
	virtual int GetCategoryIndex(int nRow);

	int GetTemplateIndex(int nRow, int nCol, bool bCheckImage = true);

	virtual bool Load(bool bReloadXML = false);	
	virtual bool Save();
	
	
	bool	OnTemplateUpdate(string strTemplate);
	void	OnActivate();
	
	bool Delete(int nRow, int nCol);
	
	bool DoEdit(int nRow, int nCol);
		
	bool SetFilterOption(bool bIsFilterOn);
	bool IsFilterOn();
	
	bool CheckByFilter(LPCSTR lpcszTemplate, int nCategoryIndex = -1, int nTemplateIndex = -1);
	
	string GetTemplateFilePath(int nRow, int nCol);	
	
	
	bool	CheckLoadImageForTempalteNone();
	
	virtual void AddTextRow(int nRow, LPCSTR lpcstrText);
	
	
	void SetIsTemplateNone(bool bIsTemplateNone);
	bool IsTemplateNone();
	
protected:
	bool	LoadImageForTempalteNone(LPCSTR lpcszImageFileName);

private:
	bool loadImage(LPCSTR lpcszFilename, int nRow, int nCol, bool bShowTemplateName = true);
	
	
	bool getImage(int nRow, int nCol, PictureHolder& pict, string* pstrFilename = NULL);
	bool setImage(int nRow, int nCol, PictureHolder pict, LPCSTR lpcszFilename = NULL, bool bShowFileName = true);
	
	bool copyImage(int nRowFrom, int nColFrom, int nRowTo, int nColTo);
	bool deleteImage(int nRow, int nCol, bool bDeleteRow);
	
	bool getPreviousImageCell(int nRow, int nCol, int& nPreviousCellRow, int& nPreviousCellCol);
	bool getNextImageCell(int nRow, int nCol, int& nNextCellRow, int& nNextCellCol, bool bAddIfNotFound = false, bool* pbIsAddedRow = NULL);
	
	void moveImageToNextCell(int nBeginRow, int nBeginCol, int nEndRow, int nEndCol);
	void moveImageToPreviousCell(int nBeginRow, int nBeginCol, int nEndRow, int nEndCol);
	
	bool isLastImageCellInCategory(int nRow, int nCol);
	void getLastImageCellInCategory(int nRow, int nCol, int& nLastCellRow, int& nLastCellCol);		
	
	bool checkAddImageRow(int nRow);
	
	void dragAndDropImage();
	void dragAndDropCategory();

	virtual int getCellWidth();	
	
	bool editCategory(int nRow);
	bool editTemplate(int nRow, int nCol);
	

	bool deleteCategory(int nRow);
	bool deleteTemplate(int nRow, int nCol);
	bool filter(LPCSTR lpcszTemplate);
	
private:
	vector<string>		m_vsTemplateFileNotInCategory;	
	
	bool				m_bIsFilterOn;
	
	Tree				m_trFilter;
	
	bool				m_bIsTemplateNone;
	
	string				m_strDefaultTemplate;
};

#define	WM_TEMPLATE_UPDATE	(WM_USER + 4321)

class TemplateLibraryDlg : public ResizeDialog
{
public:
	TemplateLibraryDlg(BOOL bMakePlot = TRUE, BOOL bIsUserDefined = TRUE, string strCategory = "", string strTemplate = "", int nPlotID = IDM_PLOT_UNKNOWN) 
		: ResizeDialog(IDD_TEMPLATE_LIBRARY_EX, "ODlg8")
	{
		m_bMakePlot = bMakePlot;

		m_strDefaultCategory = strCategory;
		m_strDefaultTemplate = strTemplate;
		m_nPlotID = nPlotID;
	}
	
	~TemplateLibraryDlg()
	{
	}	
	
	
	int Create(HWND hWndParent = NULL)
	{
		InitMsgMap();
		DWORD dwDlgOptions = 0;
		int nRet = ResizeDialog::Create(hWndParent, dwDlgOptions);
		Visible = true;
		return nRet;
	}
	
	
protected:
EVENTS_BEGIN
	ON_INIT(OnInitDialog)	
	ON_READY(OnReady)
	ON_SIZE(OnDlgResize)	
	ON_DESTROY(OnDestroy)
	ON_RESTORESIZE(OnRestoreSize)
	OnActivate(OnActivate)
	
	ON_GRID_SEL_CHANGE(IDC_PREVIEW_LIST, OnPreviewListCellSelChange)
	ON_GRID_BEFORE_EDIT(IDC_PREVIEW_LIST, OnListBeforeEdit)
	ON_GRID_BEFORE_EDIT(IDC_ICON_LIST, OnListBeforeEdit)
	ON_GRID_TOOL_BAR_CLICK(IDC_PREVIEW_LIST, OnListToolBarClick)
	ON_GRID_AFTER_EDIT(IDC_PREVIEW_LIST, OnPreviewListAfterEdit)
	ON_GRID_START_DRAG(IDC_PREVIEW_LIST, OnPreviewListStartDrag)
	ON_GRID_GIVE_FEEDBACK(IDC_PREVIEW_LIST, OnPreviewListDragFeedback)
	ON_GRID_COMPLETE_DRAG(IDC_PREVIEW_LIST, OnPreviewListDragDropFinish)
	ON_GRID_DBLCLICK(IDC_PREVIEW_LIST, OnDblClickToPlot)
	
	ON_BN_CLICKED(IDC_BUTTON_SCAN_USER_TEMPLATES, OnBtnScanUserTemplatesClick)
	ON_BN_CLICKED(IDC_BUTTON_NEW_CATEGORY, OnClickNewCategory)
	ON_BN_CLICKED(IDC_BUTTON_ADD_TEMPLATE, OnBtnAddTemplateClick)
	
	ON_BN_CLICKED(IDC_BUTTON_PLOT_SETUP, OnClickPlotSetupButton)
	ON_BN_CLICKED(IDC_BUTTON_PLOT, OnClickPlotButton)
	
	ON_BN_CLICKED(IDC_BUTTON_FILTER, OnFilterClicked)
	
	ON_GRID_MOUSE_MOVE(IDC_PREVIEW_LIST, OnGridMouseMove)
	
	ON_BN_CLICKED(IDC_BTN_ORIGINAL_DATA, OnBtnGetOriginalData)

	ON_CHANGE_PAGE( OnChangePage )
	ON_SYSCOMMAND(OnSystemCommand)
	
	ON_CHANGE_LAYER(OnActiveLayerChange)

	ON_CHANGE_SELECTION(OnSelectionChange)
	
	ON_USER_MSG(WM_TEMPLATE_UPDATE, OnTemplateUpdate)

	
	ON_HELPINFO(OnHelp)
	
EVENTS_END

	BOOL OnHelp(int &nHelpID, int nIdCtrlFocus)
	{
		nHelpID = IDD_TEMPLATE_LIBRARY_H;
		return TRUE;
	}

	BOOL    OnInitDialog();
	BOOL    OnReady();
	BOOL    OnDlgResize(int nType, int cx, int cy);
	BOOL    OnDestroy();
	BOOL    OnRestoreSize(ODWP dwSizeInfo);
	
	void    OnPreviewListCellSelChange(Control ctrl);
	
	void    OnListBeforeEdit( Control ctrl, int nRow, int nCol, BOOL* pCancel );
	void	OnListToolBarClick(Control ctrl, int nRow, int nCol, int nMouseOverIndex);
	
	void    OnPreviewListAfterEdit(Control flxControl, int nRow, int nCol);
	void    OnPreviewListStartDrag(Control ctrl, int* pnAllowEffect, int* pnRow, int* pnCol, int* pnRowSel, int* pnColSel);
	void    OnPreviewListDragFeedback(Control ctrl, int* pnAllowEffect, bool* pCancel);
	void    OnPreviewListDragDropFinish(Control ctrl, int* pnAllowEffect);
	
	BOOL    OnDblClickToPlot(Control ctrl);
	
	void    OnBtnScanUserTemplatesClick(Control ctrl);
	void    OnClickNewCategory(Control ctrl);
	void    OnBtnAddTemplateClick(Control ctrl);
	
	BOOL    OnClickPlotButton(Control ctrl);
	BOOL    OnClickPlotSetupButton(Control ctrl);
	
	BOOL    OnFilterClicked(Control ctrl);
	void    OnFilterByColMatch(Control ctrl);
	void    OnFilterByDataFrom(Control ctrl);
	
	BOOL    OnBtnGetOriginalData(Control ctrl);
	
    void    OnGridMouseMove(Control ctrl, short nButton, short nShift, float fX, float fY);
	
	void    OnActivate(UINT nState, HWND hwndOther, BOOL bMinimized);

	BOOL    OnChangePage();
	bool    OnSystemCommand(int nCmd);
	
	BOOL    OnActiveLayerChange();
	
	BOOL    OnSelectionChange();
	
	BOOL OnTemplateUpdate(WPARAM wParam, LPARAM lParam);
	
private:
	bool doSmartPlot();
	
	void setFilterControlsEnable();
	
	void updateControlWithTemplateInfo();
	
	bool isSelectedUserDefinedCategory();	
	
	void updatePreviewListByFilterOption(bool bUpdateList);
	
	string getCurrentSelectedTemplate(bool bOnlyName);
	
	void updatePlottingButtonEnableStatus();
	
	void checkErrorMessage(int nErrorType);
	
	void resizeDlg();
	void initErrorMessageControl();
	bool setErrorMessage(LPCSTR lpcstrMsg);
	
private:
	PreviewList			m_PreviewList;
	
	TemplateSettingsHelper			m_XMLHelper;
	
	bool				m_bMakePlot;	
	
	string				m_strDefaultCategory;
	string				m_strDefaultTemplate;
	int					m_nPlotID;
	
	vector<int>		    m_vnErrMsgs;
	
	BitmapRadioButton	m_btnFilter;
};

#endif //TEMPLATE_LIBRARY_DLG_H 

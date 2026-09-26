declare interface ISmartOrgChartWebPartStrings {
  // ── Property pane: pane header / groups ──
  PropertyPane_HeaderDescription: string;
  PropertyPane_Group_General: string;
  PropertyPane_Group_Branding: string;
  PropertyPane_Group_VisualStyle: string;
  PropertyPane_Group_DataSource: string;
  PropertyPane_Group_CustomAttributes: string;
  PropertyPane_Group_UserFilters: string;
  PropertyPane_Group_OrgChart: string;
  PropertyPane_Group_OrgChartFeatures: string;
  PropertyPane_Group_Directory: string;
  PropertyPane_Group_Demo: string;

  // ── Property pane: General ──
  PropertyPane_DefaultView_Label: string;

  // ── Property pane: Branding ──
  PropertyPane_AppTitle_Label: string;
  PropertyPane_AppTitle_Placeholder: string;
  PropertyPane_AppTitle_Description: string;
  PropertyPane_LogoUrl_Label: string;
  PropertyPane_LogoUrl_Placeholder: string;
  PropertyPane_LogoUrl_Description: string;
  PropertyPane_DirectoryLabel_Label: string;
  PropertyPane_OrgChartLabel_Label: string;
  PropertyPane_ViewLabel_Description: string;

  // ── Property pane: Visual Style ──
  PropertyPane_ChartTheme_Label: string;
  PropertyPane_Theme_Modern: string;
  PropertyPane_Theme_Minimal: string;
  PropertyPane_Theme_Corporate: string;
  PropertyPane_Theme_Dark: string;
  PropertyPane_Theme_Custom: string;
  PropertyPane_AccentColor_Label: string;
  PropertyPane_AccentColor_PickerAriaLabel: string;
  PropertyPane_AccentColor_HexLabel: string;
  PropertyPane_AccentColor_HexAriaLabel: string;
  PropertyPane_AccentColor_RgbLabel: string;
  PropertyPane_AccentColor_ChannelAriaLabel: string;
  PropertyPane_AccentColor_ContrastWithWhite: string;
  PropertyPane_AccentColor_PassesAA: string;
  PropertyPane_AccentColor_LargeTextOnly: string;
  PropertyPane_AccentColor_LowContrast: string;
  PropertyPane_AccentColor_DarkTextNote: string;
  PropertyPane_AccentColor_DarkerShadeNote: string;
  PropertyPane_AccentColor_InvalidHex: string;
  PropertyPane_DefaultFontSize_Label: string;
  PropertyPane_FontSize_75: string;
  PropertyPane_FontSize_85: string;
  PropertyPane_FontSize_100: string;
  PropertyPane_FontSize_115: string;
  PropertyPane_FontSize_130: string;
  PropertyPane_FontSize_150: string;
  PropertyPane_FontSize_175: string;
  PropertyPane_DefaultLayout_Label: string;
  PropertyPane_Layout_Drill: string;
  PropertyPane_Layout_Vertical: string;
  PropertyPane_Layout_Horizontal: string;

  // ── Property pane: Data Source ──
  PropertyPane_DataSource_Label: string;
  PropertyPane_DataSource_Auto: string;
  PropertyPane_DataSource_Graph: string;
  PropertyPane_DataSource_Search: string;
  PropertyPane_DataSource_Note: string;

  // ── Property pane: Custom Attributes ──
  PropertyPane_CustomAttributes_Note: string;
  PropertyPane_CustomAttributes_Heading: string;
  PropertyPane_CustomAttributes_Empty: string;
  PropertyPane_CustomAttributes_AddButton: string;
  PropertyPane_CustomAttributes_MaxTooltip: string;
  PropertyPane_CustomAttributes_FieldNameLabel: string;
  PropertyPane_CustomAttributes_FieldPlaceholder: string;
  PropertyPane_CustomAttributes_FieldError: string;
  PropertyPane_CustomAttributes_DisplayLabelLabel: string;
  PropertyPane_CustomAttributes_LabelPlaceholder: string;
  PropertyPane_CustomAttributes_RemoveTitle: string;
  PropertyPane_CustomAttributes_RemoveAriaLabel: string;
  PropertyPane_CustomAttributes_ShowInDirectory: string;
  PropertyPane_CustomAttributes_ShowInOrgChart: string;

  // ── Property pane: User Filters ──
  PropertyPane_ExcludedAccounts_Label: string;
  PropertyPane_ExcludedAccounts_Placeholder: string;
  PropertyPane_ExcludedAccounts_Description: string;
  PropertyPane_RestrictTenant_Label: string;
  PropertyPane_RestrictTenant_On: string;
  PropertyPane_RestrictTenant_Off: string;
  PropertyPane_HideGuest_Label: string;
  PropertyPane_HideGuest_On: string;
  PropertyPane_HideGuest_Off: string;
  PropertyPane_HideDisabled_Label: string;
  PropertyPane_HideDisabled_On: string;
  PropertyPane_HideDisabled_Off: string;
  PropertyPane_HideDisabled_Note: string;
  PropertyPane_HideNoJobTitle_Label: string;
  PropertyPane_HideNoJobTitle_On: string;
  PropertyPane_HideNoJobTitle_Off: string;
  PropertyPane_HideNoDept_Label: string;
  PropertyPane_HideNoDept_On: string;
  PropertyPane_HideNoDept_Off: string;

  // ── Property pane: Org Chart ──
  PropertyPane_TopLevelUser_Label: string;
  PropertyPane_TopLevelUser_Placeholder: string;
  PropertyPane_TopLevelUser_Description: string;
  PropertyPane_LevelsBelow_Label: string;
  PropertyPane_DottedLine_Label: string;
  PropertyPane_DottedLine_Placeholder: string;
  PropertyPane_DottedLine_Description: string;
  PropertyPane_DefaultZoom_Label: string;
  PropertyPane_Zoom_Auto: string;
  PropertyPane_Zoom_50: string;
  PropertyPane_Zoom_75: string;
  PropertyPane_Zoom_100: string;
  PropertyPane_Zoom_125: string;
  PropertyPane_Zoom_150: string;

  // ── Property pane: Org Chart Features ──
  PropertyPane_FeatureToggles_Note: string;
  PropertyPane_FindMe_Label: string;
  PropertyPane_LayoutToggle_Label: string;
  PropertyPane_Stats_Label: string;
  PropertyPane_DeptFilter_Label: string;
  PropertyPane_UserFilter_Label: string;
  PropertyPane_Visible: string;
  PropertyPane_Hidden: string;

  // ── Property pane: Directory ──
  PropertyPane_PageSize_Label: string;
  PropertyPane_PageSize_AriaLabel: string;
  PropertyPane_PageSize_Note: string;

  // ── Property pane: Demo ──
  PropertyPane_UseDemoData_Label: string;
  PropertyPane_UseDemoData_On: string;
  PropertyPane_UseDemoData_Off: string;

  // ── Presence status labels ──
  Presence_Available: string;
  Presence_Busy: string;
  Presence_DoNotDisturb: string;
  Presence_BeRightBack: string;
  Presence_Away: string;
  Presence_Offline: string;

  // ── Header / top bar (SmartOrgChart.tsx) ──
  Header_DirectoryLabel: string;
  Header_OrgChartLabel: string;
  Header_SwitchToView: string;
  Header_CompanyLogoAlt: string;
  Header_RefreshingAria: string;
  Header_RefreshAria: string;
  Header_RefreshTitle: string;
  Header_NotLoadedYet: string;
  Header_JustNow: string;
  Header_MinutesAgo: string;
  Header_HoursAgo: string;
  Header_PreferencesTitle: string;
  Header_PreferencesAria: string;

  // ── Employee Directory ──
  Directory_LoadingLabel: string;
  Directory_LoadFailedFallback: string;
  Directory_LoadFailedPrefix: string;
  Directory_AllDepartments: string;
  Directory_AllOffices: string;
  Directory_StatusDisabled: string;
  Directory_StatusGuest: string;
  Directory_PresenceStatus: string;
  Directory_ChatInTeams: string;
  Directory_ChatWithInTeams: string;
  Directory_ManagesCount: string;
  Directory_ManagesCountWithTotal: string;
  Directory_ReportsCountWithTotal: string;
  Directory_ColumnName: string;
  Directory_ColumnJobTitle: string;
  Directory_ColumnDepartment: string;
  Directory_ColumnOffice: string;
  Directory_ColumnEmail: string;
  Directory_ColumnPhone: string;
  Directory_ColumnReports: string;
  Directory_ColumnChat: string;
  Directory_PrevPage: string;
  Directory_NextPage: string;
  Directory_PreviousPageAria: string;
  Directory_NextPageAria: string;
  Directory_SearchPlaceholder: string;
  Directory_SearchAria: string;
  Directory_FilterByDepartmentAria: string;
  Directory_FilterByOfficeAria: string;
  Directory_ClearSearchTitle: string;
  Directory_ClearButton: string;
  Directory_ClearFiltersButton: string;
  Directory_ExportCsvTitle: string;
  Directory_ExportCsvButton: string;
  Directory_ViewModeAria: string;
  Directory_CardViewTitle: string;
  Directory_ListViewTitle: string;
  Directory_AlphabetFilterAria: string;
  Directory_ShowAllTitle: string;
  Directory_OtherLettersTitle: string;
  Directory_FilterByLetterTitle: string;
  Directory_ResultCountSingular: string;
  Directory_ResultCountPlural: string;
  Directory_PageOf: string;
  Directory_FilterActiveSingular: string;
  Directory_FilterActivePlural: string;
  Directory_NoResults: string;
  Directory_PaginationAria: string;

  // ── Org chart toolbar / shared chart chrome ──
  Chart_Layout_Drill: string;
  Chart_Layout_Vertical: string;
  Chart_Layout_Horizontal: string;
  Chart_SearchPlaceholder: string;
  Chart_SearchAria: string;
  Chart_MatchSingular: string;
  Chart_MatchPlural: string;
  Chart_ResultsListAria: string;
  Chart_ExpandAll: string;
  Chart_CollapseAll: string;
  Chart_Expanding: string;
  Chart_Loading: string;
  Chart_FindMeTitle: string;
  Chart_ViewLayoutButton: string;
  Chart_ViewLayoutTitle: string;
  Chart_ViewLayoutPanelTitle: string;
  Chart_StatsSummaryTitle: string;
  Chart_FilterByDepartmentTitle: string;
  Chart_FilterByDepartmentWithCountAria: string;
  Chart_FilterByDepartmentPanelTitle: string;
  Chart_ClearAllFilters: string;
  Chart_FilterUserTypesTitle: string;
  Chart_DownloadPdfTitle: string;
  Chart_DownloadCsvTitle: string;
  Chart_ResetRootTitle: string;
  Chart_ViewFromPersonPlaceholder: string;
  Chart_ViewFromPersonAria: string;
  Chart_ZoomOutTitle: string;
  Chart_ZoomInTitle: string;
  Chart_ZoomResetTitle: string;
  Chart_ShowInChartAria: string;
  Chart_ShowInChartTitle: string;
  Chart_FilterMembers: string;
  Chart_FilterGuests: string;
  Chart_FilterDisabled: string;
  Chart_SetupTitle: string;
  Chart_SetupSubtitle: string;
  Chart_SetupPlaceholder: string;
  Chart_SetupInputAria: string;
  Chart_SetupLoadButton: string;
  Chart_SetupHint: string;
  Chart_BuildingLabel: string;
  Chart_RetryButton: string;
  Chart_NotFoundFromSetup: string;
  Chart_NotFoundWithSettings: string;
  Chart_LoadFailedGeneric: string;
  Chart_LoadFailedPermissions: string;
  Chart_LoadFailedForPerson: string;
  Chart_FindMeNotFound: string;
  Chart_StatPeople: string;
  Chart_StatMembers: string;
  Chart_StatGuests: string;
  Chart_StatDepts: string;
  Chart_ReportingLineAria: string;
  Chart_BackToFullOrgTitle: string;
  Chart_FullOrgLabel: string;
  Chart_FocusOnPerson: string;
  Chart_OrgChartAria: string;
  Chart_ExportNote_DrillLevel: string;
  Chart_ExportNote_PartialLevels: string;
  Chart_AllowPopupsForPdf: string;

  // ── Person card ──
  PersonCard_ProfileAria: string;
  PersonCard_CloseTitle: string;
  PersonCard_CloseAria: string;
  PersonCard_StatusDisabled: string;
  PersonCard_StatusGuest: string;
  PersonCard_DirectReportSingular: string;
  PersonCard_DirectReportPlural: string;
  PersonCard_TotalSuffix: string;
  PersonCard_CopyEmailTitle: string;
  PersonCard_EmailCopiedAria: string;
  PersonCard_ReportsTo: string;
  PersonCard_FocusOnPerson: string;
  PersonCard_FocusOrgChartOnPerson: string;
  PersonCard_DottedLine: string;
  PersonCard_DottedLineManager: string;
  PersonCard_DottedLineReport: string;
  PersonCard_ChatButton: string;
  PersonCard_EmailButton: string;
  PersonCard_FocusButton: string;
  PersonCard_FocusTitle: string;

  // ── Org node card ──
  OrgNode_ViewProfile: string;
  OrgNode_FocusOrgChartOn: string;
  OrgNode_Collapse: string;
  OrgNode_Expand: string;
  OrgNode_DirectReportsSuffix: string;
  OrgNode_StatusDisabled: string;
  OrgNode_StatusGuest: string;
  OrgNode_TotalReportsAcrossLevels: string;

  // ── Drill view ──
  Drill_ReportingLineAria: string;
  Drill_BackToTopTitle: string;
  Drill_BackToTopAria: string;
  Drill_GoBackTo: string;
  Drill_ViewProfileTitle: string;
  Drill_ViewProfileAria: string;
  Drill_ProfileButton: string;
  Drill_LoadingLabel: string;
  Drill_DirectReports: string;
  Drill_DirectReportsOfAria: string;
  Drill_DirectReportsAria: string;
  Drill_ViewPersonProfile: string;
  Drill_ShowPersonReports: string;
  Drill_NoDirectReports: string;

  // ── Settings panel ──
  Settings_AlphabetFirstName: string;
  Settings_AlphabetLastName: string;
  Settings_CardSizeSmall: string;
  Settings_CardSizeMedium: string;
  Settings_CardSizeLarge: string;
  Settings_HeaderText: string;
  Settings_SaveButton: string;
  Settings_CancelButton: string;
  Settings_DiscardButton: string;
  Settings_SectionDirectory: string;
  Settings_AlphabetFilterByLabel: string;
  Settings_CardSizeLabel: string;
  Settings_FontSizeLabel: string;
  Settings_SectionCardsAndFields: string;
  Settings_ShowOnCardsLabel: string;
  Settings_EmailAddressLabel: string;
  Settings_PhoneNumberLabel: string;
  Settings_DepartmentLabel: string;
  Settings_OfficeLocationLabel: string;
  Settings_SectionOrgChart: string;
  Settings_ManagerLevelsLabel: string;
  Settings_CompactCardsLabel: string;
  Settings_CompactCardsOn: string;
  Settings_CompactCardsOff: string;
  Settings_Hint: string;
  Settings_SectionDemoData: string;
  Settings_DatasetSizeLabel: string;
  Settings_PeopleCountSuffix: string;

  // ── Export (CSV / PDF) ──
  Export_ColumnName: string;
  Export_ColumnJobTitle: string;
  Export_ColumnDepartment: string;
  Export_ColumnOffice: string;
  Export_ColumnEmail: string;
  Export_ColumnPhone: string;
  Export_ColumnManager: string;
  Export_ColumnLevel: string;
  Export_PdfDocumentTitle: string;
  Export_PdfHeading: string;
  Export_PdfRootLabel: string;
}

declare module 'SmartOrgChartWebPartStrings' {
  const strings: ISmartOrgChartWebPartStrings;
  export = strings;
}

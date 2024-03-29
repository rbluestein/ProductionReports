if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FN_PATCOUNT]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FN_PATCOUNT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FN_Repetitive_Str_Parse]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FN_Repetitive_Str_Parse]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SalesByStore]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[SalesByStore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[udf_Convert2TitleCase]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[udf_Convert2TitleCase]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_BuildProductSegment3_Extended]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_BuildProductSegment3_Extended]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_BuildProductSegment3_Standard]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_BuildProductSegment3_Standard]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_BuildProductSegment3_Standard_Horizontal]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_BuildProductSegment3_Standard_Horizontal]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_DateCompare]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_DateCompare]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_DatetimeToDate]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_DatetimeToDate]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_ESTToLocal]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_ESTToLocal]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_Extract_Substring_From_String]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_Extract_Substring_From_String]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetAppsAndPolsDataJoin]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetAppsAndPolsDataJoin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetBVIAppStatus]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetBVIAppStatus]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetClientJoin]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetClientJoin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetClientWhere]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetClientWhere]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetClientWhereSelect]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetClientWhereSelect]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetDaysInMonth]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetDaysInMonth]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetEnrolled]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetEnrolled]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetPreferredAPSRelationCode]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetPreferredAPSRelationCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetProductFieldsTable]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetProductFieldsTable]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetSqlBoilerplate]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetSqlBoilerplate]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetTableFromList]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetTableFromList]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_GetTier]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_GetTier]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_IsAPSLegal]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_IsAPSLegal]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_IsDateBetween]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_IsDateBetween]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_IsDateEqual]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_IsDateEqual]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_IsGuid]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_IsGuid]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_IsTestID]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_IsTestID]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_LocalToEST]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_LocalToEST]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_PadLeft]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_PadLeft]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_PadRight]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_PadRight]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_PassIAMSLinkTest]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_PassIAMSLinkTest]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_RptBuildProductSegment3]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_RptBuildProductSegment3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_SalesByStore]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_SalesByStore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_Split]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_Split]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_StringCleanser]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_StringCleanser]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_ToInt]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_ToInt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_ToProper]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_ToProper]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufn_get_FirstDayOfFollowingMonth]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufn_get_FirstDayOfFollowingMonth]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spDbDefrag]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spDbDefrag]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spDbReIndex]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spDbReIndex]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spRunReplication]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spRunReplication]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spWhere_Am_I]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spWhere_Am_I]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_CCM2010BoardRecords]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_CCM2010BoardRecords]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_CCM2010BoardRecordsThisEmp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_CCM2010BoardRecordsThisEmp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_CallHistoryWorklistSubTableQueryAdj]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_CallHistoryWorklistSubTableQueryAdj]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_DeleteCall]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_DeleteCall]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_DeleteCallMonitor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_DeleteCallMonitor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_DeleteIAMSCancelled]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_DeleteIAMSCancelled]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_DeleteProduct]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_DeleteProduct]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_DeleteUndeleteCallMonitor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_DeleteUndeleteCallMonitor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GetConditionedFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GetConditionedFields]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GetDailyHoursWorked]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GetDailyHoursWorked]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GetDatesUpdateRequired]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GetDatesUpdateRequired]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GetExcelData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GetExcelData]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_JonSnavelyDiopitt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_JonSnavelyDiopitt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_JonSnavelyGNC]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_JonSnavelyGNC]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ProjectReportsMigrate_APSDataChanged]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ProjectReportsMigrate_APSDataChanged]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ProjectReportsMigrate_NoAPSRecord]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ProjectReportsMigrate_NoAPSRecord]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ProjectReportsMigrate_NoAPSRecord_Mike]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ProjectReportsMigrate_NoAPSRecord_Mike]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_QADC_Verify_EmpTransmittal_vs_ClientCallActivity]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_QADC_Verify_EmpTransmittal_vs_ClientCallActivity]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildCalendar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildCalendar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildClientSegment1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildClientSegment1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildClientSegment101]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildClientSegment101]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildClientSegment2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildClientSegment2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildProductCompactor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildProductCompactor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildProductSegmentExtended1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildProductSegmentExtended1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildProductSegmentExtended2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildProductSegmentExtended2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildProductSegmentStandard1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildProductSegmentStandard1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildProductSegmentStandard2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildProductSegmentStandard2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildProductSegments]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildProductSegments]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildProductSegmentsThisClient]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildProductSegmentsThisClient]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildStart]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildStart]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptBuildTempDatesUpdateRequired]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptBuildTempDatesUpdateRequired]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptStartMemOnly]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptStartMemOnly]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RptUpdateData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RptUpdateData]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Rpt_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Rpt_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Rpt_All_102011]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Rpt_All_102011]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Rpt_All_NameOnly]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Rpt_All_NameOnly]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Rpt_All_NoManHours]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Rpt_All_NoManHours]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Rpt_All_Old]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Rpt_All_Old]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Rpt_EnrollerProductivity]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Rpt_EnrollerProductivity]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempBobErrorHandling]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempBobErrorHandling]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempCursor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempCursor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempCursorCallSP2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempCursorCallSP2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempDBExist]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempDBExist]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempGetTblDef]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempGetTblDef]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempIterateMemTableGood]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempIterateMemTableGood]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempReturnTableToSP1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempReturnTableToSP1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempReturnTableToSP2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempReturnTableToSP2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempRptStartPhysicalTable]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempRptStartPhysicalTable]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempSortTable]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempSortTable]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_TempSplitWithCursor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_TempSplitWithCursor]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE FUNCTION FN_PATCOUNT 
/*****************************************************************
** Name        	: 	FN_PATCOUNT
**
** Description 	: 	Will return a count of how many times the search
**			            pattern occurs in the submitted string. 
**			            (A custom Metadata function)
**			            
** Written By	  : 	
** Uses Tables	:	  
** Parameters  	: 	
** Returns     	: 	
** Modifications:	            
**                  
*****************************************************************/
(       @PATX      VARCHAR(255),
        @STR       VARCHAR(8000)
)
RETURNS SMALLINT
AS
BEGIN
 
        
DECLARE @PATCOUNT  SMALLINT,
        @PATIDX    SMALLINT,
        @PATLEN    TINYINT
 
SELECT  @PATCOUNT = 1,
        @PATIDX = 1,
        @PATLEN = LEN(@PATX)
 
WHILE @PATIDX <= LEN(@STR)
BEGIN
 IF (SELECT SUBSTRING(@STR, @PATIDX - @PATLEN, @PATLEN)) = @PATX
  SET @PATCOUNT = @PATCOUNT + 1
  SET @PATIDX = @PATIDX + 1
END
 
RETURN @PATCOUNT
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE FUNCTION dbo.FN_Repetitive_Str_Parse 
/*****************************************************************
** Name        	: 	dbo.FN_Repetitive_Str_Parse
**
** Description 	: 	DECLARE @KEY_DATA VARCHAR(200),
**                  @VAL_DATA VARCHAR(200)
**          SELECT  @KEY_DATA = 'AB12345~PA~ORD0001',
**                  @VAL_DATA = '253~USER1'
**          SELECT  dbo.FN_Repetitive_Str_Parse (@KEY_DATA, '~', 1) AS PART,
**                  dbo.FN_Repetitive_Str_Parse (@KEY_DATA, '~', 2) AS LOC,
**                  dbo.FN_Repetitive_Str_Parse (@KEY_DATA, '~', 3) AS [ORDER],
**                  dbo.FN_Repetitive_Str_Parse (@VAL_DATA, '~', 1) AS SHIP_QTY,
**                  dbo.FN_Repetitive_Str_Parse (@VAL_DATA, '~', 2) AS NETID,
**                  GETDATE() AS DATE_STAMP
**          
			            
** Written By	  : 	
** Uses Tables	:	  
** Parameters  	: 	
** Returns     	: 	
** Modifications:	  
**                  
*****************************************************************/
(        @STR2PARSE           VARCHAR(5000), 
         @CHECK_CHAR          VARCHAR(10),
         @DATA_ELEMENT        SMALLINT
)
RETURNS VARCHAR(100)
 
AS
 
BEGIN
DECLARE @STR_VAL       VARCHAR(100),
        @STR_INC       SMALLINT,
        @STR_LEN       SMALLINT,
        @FIELD_CNT     TINYINT,
        @CHK_LEN       TINYINT,
        @MAX_FIELD     TINYINT,
        @FIELD_IDX     SMALLINT,
        @VAL_INDEX     SMALLINT,
        @TEST_CHAR     VARCHAR(10),
        @TEST_STRING1  VARCHAR(100),
        @RETURN_STAT   TINYINT
 
 
---===<<<  Create temp table that contains Positional Data  >>>===---
 
DECLARE @POS_DATA TABLE
( Element_type     CHAR(4),
  Element_Count    TINYINT,
  Element_Position SMALLINT
)
DECLARE @VAL_LIST TABLE
( STR_VALUE        VARCHAR(800),
  DATA_ELEMENT     SMALLINT
)
 
 
---===<<<  Var Initilalization  >>>===---
 
 
SELECT @STR_LEN = LEN(@STR2PARSE),
       @STR_INC = 0,
       @FIELD_CNT = 0,
       @CHK_LEN = LEN(@CHECK_CHAR),
       @TEST_CHAR = ''
 
---===<<<  Parse string   >>>===---
-- ****************************************
-- Generate a table recording the position in the string
-- where the delimiter exists.
-- ****************************************
WHILE @STR_INC <= @STR_LEN
  BEGIN
    IF SUBSTRING(@STR2PARSE, @STR_INC - @CHK_LEN, @CHK_LEN) = @CHECK_CHAR
    BEGIN
      SET @FIELD_CNT = @FIELD_CNT + 1
      INSERT INTO @POS_DATA
      VALUES ('LIST', @FIELD_CNT, @STR_INC)
    END
    SET @STR_INC = @STR_INC + 1
  END 
 
-- ****************************************
-- Test to see if the string ended with the delimiter.
-- If it did not, Record the position of the last field.
-- ****************************************
IF (SELECT SUBSTRING(@STR2PARSE, MAX(ELEMENT_POSITION) + 1, 1)
              FROM @POS_DATA) <> ''
  BEGIN
    SET @FIELD_CNT = @FIELD_CNT + 1
    INSERT INTO @POS_DATA
    VALUES ('LIST', @FIELD_CNT, LEN(@STR2PARSE) + 1)
  END
 
 
 
---===<<<  Parse values and Columns  >>>===---
SELECT  @FIELD_CNT = MIN(ELEMENT_COUNT),
        @MAX_FIELD = MAX(ELEMENT_COUNT)
  FROM @POS_DATA
 
WHILE @FIELD_CNT <= @MAX_FIELD
BEGIN
  SELECT @FIELD_IDX = CASE 
      WHEN @FIELD_CNT = 1 THEN 1
      ELSE @STR_INC - @CHK_LEN + 1
    END
 
  SELECT @STR_INC = ELEMENT_POSITION FROM @POS_DATA
      WHERE ELEMENT_TYPE = 'LiST'
      AND   ELEMENT_COUNT = @FIELD_CNT
 
  IF @FIELD_CNT = @MAX_FIELD
    SET @TEST_STRING1 = LTRIM(SUBSTRING(@STR2PARSE, @FIELD_IDX, @STR_INC - @FIELD_IDX))
  ELSE
    SET @TEST_STRING1 = LTRIM(SUBSTRING(@STR2PARSE, @FIELD_IDX, @STR_INC - @FIELD_IDX - @CHK_LEN))
 
    INSERT INTO @VAL_LIST
    SELECT @TEST_STRING1,
           @FIELD_CNT
 
  SELECT @FIELD_CNT = @FIELD_CNT + 1
 
END
 
-- **********************************************************
-- Fetch the value requested by user.
-- If necessary clean out any occurrences of the delimiter
-- **********************************************************
SELECT @STR_VAL = CASE
          WHEN CHARINDEX(@CHECK_CHAR, STR_VALUE, 2) > 0 THEN LEFT(STR_VALUE, LEN(STR_VALUE) - @CHK_LEN)
          ELSE STR_VALUE 
        END
  FROM @VAL_LIST
  WHERE DATA_ELEMENT = @DATA_ELEMENT  
 
 
RETURN ISNULL(@STR_VAL, '')
 
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO







CREATE FUNCTION SalesByStore (@storeid varchar(30))
   RETURNS @t TABLE (title varchar(80) NOT NULL,
                     qty   smallint    NOT NULL)  AS
BEGIN
   INSERT @t (title, qty)
      SELECT t.title, s.qty
      FROM   sales s
      JOIN   titles t ON t.title_id = s.title_id
      WHERE  s.stor_id  = @storeid
   RETURN
END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  FUNCTION udf_Convert2TitleCase (@InputString varchar(4000) )
RETURNS VARCHAR(4000)
AS
/*	Test Function
	SELECT dbo.udf_Convert2TitleCase('This function will convert this string to title ')
*/
BEGIN
DECLARE @Index          INT
DECLARE @Char           CHAR(1)
DECLARE @OutputString   VARCHAR(255)
SET @OutputString = LOWER(@InputString)
SET @Index = 2
SET @OutputString =
   STUFF(@OutputString, 1, 1,UPPER(SUBSTRING(@InputString,1,1)))
WHILE @Index <= LEN(@InputString)
BEGIN
SET @Char = SUBSTRING(@InputString, @Index, 1)
IF @Char IN (' ', ';', ':', '!', '?', ',', '.', '_', '-', '/', '&','''', '(')
IF @Index + 1 <= LEN(@InputString)
BEGIN
IF @Char != '''' OR
UPPER(SUBSTRING(@InputString, @Index + 1, 1)) != 'S'
SET @OutputString =
   STUFF(@OutputString, @Index + 1, 1,UPPER(SUBSTRING(@InputString, @Index + 1, 1)))
END
SET @Index = @Index + 1
END
RETURN ISNULL(@OutputString,'')
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE      FUNCTION dbo.ufn_BuildProductSegment3_Extended
(
 @ClientID varchar(20),
 @ProductID varchar(20),
 @Selection varchar(200),
@TestedDate varchar(20),
@CurDate smalldatetime,
@AddlCondition varchar(200),
@Prefix varchar(10)
)
RETURNS varchar(1000)
AS
BEGIN
	DECLARE @Results varchar(1000)
	SELECT @Results = 'SELECT ' + rtrim(@Selection) + ' FROM Alt_ProductData alt '
	SELECT @Results = @Results + 'INNER JOIN EmpProductTransmittal ept on alt.AltProductDataID = ept.AltProductDataID '
	SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpTransmittal et2 on ept.ActivityID = et2.ActivityID '
	SELECT @Results = @Results + 'INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID '
	SELECT @Results = @Results + dbo.ufn_GetClientJoin(@ClientID, @Prefix)
	SELECT @Results = @Results + 'WHERE ept.LogicalDelete = 0 AND '
	SELECT @Results = @Results + dbo.ufn_GetClientWhere(@ClientID, @Prefix)
	SELECT @Results = @Results + ' et2.EnrollerID = u.UserID AND '

	-- 11/27/2009
	SELECT @Results = @Results + 'dbo.ufn_IsAPSLegal(ept.ActivityID, ept.AppID, ept.ClientID, ept.ProductID) = 1 AND '

	SELECT @Results = @Results + 'ept.ProductID = ''' + @ProductID + ''' AND '

	IF @AddlCondition <> ''
		SELECT @Results = @Results + @AddlCondition + ' AND '

	--SELECT @Results = @Results + 'dbo.ufn_IsDateEqual(@TestedDate, @CurDate) = 1 AND et2.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Results = @Results + 'dbo.ufn_IsDateEqual(''' + Cast(@TestedDate as varchar(20)) + ''', ''' + Cast(@CurDate as varchar(20)) + ''') = 1 AND et2.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Results = @Results  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '

	RETURN @Results
END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO









CREATE         FUNCTION dbo.ufn_BuildProductSegment3_Standard
(
@ClientID varchar(20),
@ProductID varchar(20),
@Selection varchar(200),
@TestedDate varchar(20),
@CurDate smalldatetime,
@AddlCondition varchar(200),
@Prefix varchar(10)
)
RETURNS varchar(1000)
AS
BEGIN
	DECLARE @Results varchar(1000)
	SELECT @Results = 'SELECT ' + rtrim(@Selection) + ' FROM IAMS..AppsAndPolsSummary aps '
	SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpTransmittal et2 on aps.ActivityID = et2.ActivityID '
	
	-- // 1/25/2011
	--SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpProductTransmittal ept on aps.ActivityID = ept.ActivityID AND aps.AppID = ept.AppID '
	SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpProductTransmittal ept on aps.AppID = ept.AppID '

	SELECT @Results = @Results + 'INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID '
	SELECT @Results = @Results + dbo.ufn_GetClientJoin(@ClientID, @Prefix)
	SELECT @Results = @Results + 'WHERE ept.LogicalDelete = 0 AND '
	SELECT @Results = @Results + dbo.ufn_GetClientWhere(@ClientID, @Prefix)
	SELECT @Results = @Results + ' aps.LicensedEnroller = u.UserID AND '
	--SELECT @Results = @Results + 'aps.ProductID = ' + @ProductID + 'AND '
	SELECT @Results = @Results + 'aps.ProductID = ''' + @ProductID + ''' AND '
	SELECT @Results = @Results + ' AND aps.BVIAppStatus <> ''CANCELLED'' AND '

	IF @AddlCondition <> ''
		SELECT @Results = @Results + @AddlCondition + ' AND '

	--SELECT @Results = @Results + 'dbo.ufn_IsDateEqual(@TestedDate, @CurDate) = 1 AND et2.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Results = @Results + 'dbo.ufn_IsDateEqual(''' + Cast(@TestedDate as varchar(20)) + ''', ''' + Cast(@CurDate as varchar(20)) + ''') = 1 AND et2.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Results = @Results  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '

	RETURN @Results
END









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE        FUNCTION dbo.ufn_BuildProductSegment3_Standard_Horizontal
(
@ClientID varchar(20),
@ProductID varchar(20),
@TestedDate varchar(20),
@CurDate smalldatetime,
@AddlCondition varchar(200),
@Prefix varchar(10)
)
RETURNS varchar(1000)
AS
BEGIN
	DECLARE @Results varchar(1000)
	SELECT @Results = ' FROM IAMS..AppsAndPolsSummary aps '
	SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpTransmittal et2 on aps.ActivityID = et2.ActivityID '
	SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpProductTransmittal ept on aps.ActivityID = ept.ActivityID AND aps.AppID = ept.AppID '
	SELECT @Results = @Results + 'INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID '
	SELECT @Results = @Results + dbo.ufn_GetClientJoin(@ClientID, @Prefix)
	SELECT @Results = @Results + 'WHERE ept.LogicalDelete = 0 AND '
	SELECT @Results = @Results + dbo.ufn_GetClientWhere(@ClientID, @Prefix)
	SELECT @Results = @Results + ' aps.LicensedEnroller = u.UserID AND '
	SELECT @Results = @Results + 'aps.ProductID = ''' + @ProductID + ''' AND '
	SELECT @Results = @Results + ' AND aps.BVIAppStatus <> ''CANCELLED'' AND '

	IF @AddlCondition <> ''
		SELECT @Results = @Results + @AddlCondition + ' AND '

	SELECT @Results = @Results + 'dbo.ufn_IsDateEqual(''' + Cast(@TestedDate as varchar(20)) + ''', ''' + Cast(@CurDate as varchar(20)) + ''') = 1 AND et2.SupervisorApprovalDate IS NOT NULL AND '
	--SELECT @Results = @Results + 'dbo.ufn_IsDateEqual(@TestedDate, @CurDate) = 1 AND et2.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Results = @Results  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '

	RETURN @Results
END








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


/*
SELECT Results = dbo.ufn_DateCompare('1/1/2009', '1/2/2009', 1)   -- Return value: -1
SELECT Results = dbo.ufn_DateCompare('1/2/2009', '1/2/2009', 1)   -- Return value: 0
SELECT Results = dbo.ufn_DateCompare('1/3/2009', '1/2/2009', 1)   -- Return value: 1
*/


--/*
CREATE     FUNCTION ufn_DateCompare
(
	@Date1 datetime,
	@Date2 datetime,
	@CompareDatePartOnly bit
)

RETURNS int

AS

--*/
BEGIN

/*	
	DECLARE @Date1 datetime
	DECLARE @Date2 datetime
	DECLARE @CompareDatePartOnly bit

	Select @Date1 = '2009-10-09 22:44:13.031'
	Select @Date2 = '2009-10-09 22:44:13.021'
	SELECT @CompareDatePartOnly  = 0
*/

	DECLARE @Results int
	DECLARE @Date1Var varchar(17)
	DECLARE @Date2Var varchar(17)

	SELECT @Date1Var = 
	Cast(Datepart(YEAR, @Date1)as varchar(4)) +
	Replace(STR(Cast(Datepart(month, @Date1)as varchar(2)), 2), ' ', '0') +
	Replace(STR(Cast(Datepart(day, @Date1)as varchar(2)), 2), ' ', '0')

	If @CompareDatePartOnly = 0
	Begin
		SELECT @Date1Var = @Date1Var +
		Replace(STR(Cast(Datepart(hour, @Date1)as varchar(2)), 2), ' ', '0') +
		Replace(STR(Cast(Datepart(minute, @Date1)as varchar(2)), 2), ' ', '0') +
		Replace(STR(Cast(Datepart(second, @Date1)as varchar(2)), 2), ' ', '0') +
		--Replace(STR(Cast(Datepart(millisecond, @Date1)as varchar(3)), 3), ' ', '0')
		dbo.ufn_PadLeft(Cast(DatePart(millisecond, @Date1) as varchar(3)), 3, '0')
	End

	SELECT @Date2Var = 
	Cast(Datepart(YEAR, @Date2)as varchar(4)) +
	Replace(STR(Cast(Datepart(month, @Date2)as varchar(2)), 2), ' ', '0') +
	Replace(STR(Cast(Datepart(day, @Date2)as varchar(2)), 2), ' ', '0')

	If @CompareDatePartOnly = 0
	Begin
		SELECT @Date2Var = @Date2Var +
		Replace(STR(Cast(Datepart(hour, @Date2)as varchar(2)), 2), ' ', '0') +
		Replace(STR(Cast(Datepart(minute, @Date2)as varchar(2)), 2), ' ', '0') +
		Replace(STR(Cast(Datepart(second, @Date2)as varchar(2)), 2), ' ', '0') +
		--Replace(STR(Cast(Datepart(millisecond, @Date2)as varchar(3)), 3), ' ', '0') 
		dbo.ufn_PadLeft(Cast(DatePart(millisecond, @Date2) as varchar(3)), 3, '0')
	End

	If @Date1Var = @Date2Var
		SELECT @Results = 0
	Else If @Date1Var < @Date2Var
		Select @Results = -1
	Else
		Select @Results = 1

	Return @Results
--Select @Results
--select @Date1Var, @Date2Var
--select dbo.ufn_PadLeft(Cast(DatePart(millisecond, @Date1) as varchar(3)), 3, '0')
--select DatePart(millisecond, @Date1)

END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE  FUNCTION dbo.ufn_DatetimeToDate 
      (@DateIn as DateTime) 
 RETURNS varchar(10) 
 AS 
 BEGIN 
       
 declare @month varchar(4)  
 declare @day varchar(4)  
 declare @year varchar(4)  
 declare @retval varchar(10) 
 set @month = datepart(m, @DateIn)  
 set @month = right('00' + @month, 2)  
 set @day = datepart(d, @DateIn)  
 set @day = right('00' + @day, 2)  
 set @year = datepart(yyyy, convert(varchar(8),@DateIn, 112))  
 set @retval = @month + '/' + @day + '/' + @year  
       
      RETURN @retval 
 END 
 



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE  FUNCTION ufn_ESTToLocal
(
	@ESTTime datetime,
	@LocalGreenwichOffset int
)

RETURNS datetime

AS


BEGIN


/*
	DECLARE @ESTTime datetime
	DECLARE @LocalGreenwichOffset int
	SELECT @ESTTime  = '6/25/2009 2:06:00 PM'
	SELECT @LocalGreenwichOffset  = 7
*/

	DECLARE @Results datetime
	DECLARE @TotalOffset int
	DECLARE @GreenwichESTOffset int
	SELECT @GreenwichESTOffset = 5
	SET @TotalOffset = @GreenwichESTOffset - @LocalGreenwichOffset


	IF @ESTTime IS NULL 
	Begin
		SELECT @Results = NULL
	End
	Else
	Begin
		Select @Results = DATEADD(hour, @TotalOffset, @ESTTime)
	End
	Return @Results
	--Select Results = @Results


END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE function dbo.ufn_Extract_Substring_From_String
		(@vc_CharString as varchar(3000))
returns varchar(3000)
as
/*
Name: ufn_ScrubNonNumeric 
Description: Only grab Enroller Name from the pkey no matter the length inside the string

Notes: When this returns a value, it keeps the leading and ending spaces. So, how to get around that is to use selection criteria:

	SELECT substring(dbo.ufn_Extract_Substring_From_String(FieldName),24,14)

Created: 12/13/2007
Created By: Ken Brown
--------------------------------------------------------
Changelog
*/
begin 
declare @sub varchar(1)
declare @retval varchar(3000)
declare @strLen int
declare @idx as int
set @strlen = len(@vc_CharString)
set @idx = 1
set @retval = ''
while (@idx < (@strLen+1))
	begin
	set @sub = substring(@vc_CharString, @idx, 1)
	if (ascii(@sub) not in (ascii(''), ascii('.'), ascii('0'), ascii('1'), ascii('2'), ascii('3'), ascii('4'), ascii('5'), ascii('6'), ascii('7'), ascii('8'), ascii('9')))
		begin
		set @retval = @retval + cast(@sub as varchar(1))
		end
	else
		set @retval = @retval + ' ' 
		
	set @idx = @idx + 1
	end
return @retval
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  FUNCTION dbo.ufn_GetAppsAndPolsDataJoin()

RETURNS varchar(400)
AS
BEGIN
	DECLARE @Results varchar(400)

	SELECT @Results = 'LEFT JOIN IAMS..AppsAndPolsData apd ON aps.AppID = apd.AppID '
	SELECT @Results = @Results + 'AND '
	SELECT @Results = @Results + '('
	SELECT @Results = @Results + '(aps.ProductID IN (''TRANSUL'', ''TMARKCOMBO'')  AND apd.FieldName=''EZValue'' and apd.FieldData = ''true'') '
	SELECT @Results = @Results + 'OR '
	SELECT @Results = @Results + '(aps.ProductID = ''TMARKUL''  AND CharIndex(''EZV'', apd.FieldData) > 0) '
	SELECT @Results = @Results + 'OR '
	SELECT @Results = @Results + '(aps.ProductID = ''ALLSTATEUL'' AND apd.FieldName = ''Future Purchase Option Rider'') '
	SELECT @Results = @Results + ') '

	RETURN @Results
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



--   Select Results = dbo.ufn_IsGuid('3981bc36-c5f5-458a-a8ef-5afffb0003ee')
--   Select Results = dbo.ufn_IsGuid('')
--   Select Results = dbo.ufn_IsGuid(null)

--   Select Results = dbo.ufn_GetBVIAppStatus('7322bec5-06c3-44f3-a50a-cd5b05fc7506', '7322bec5-06c3-44f3-a50a-cd5b05fc7506', '20091118134835310')


-- // Rule for call
-- // Call is legal whether or not IAMS call record exists.
-- // Convert-to-call calls have no bearing on this rule.

-- // Rule for product
-- // Product is legal only when it is possible to link the EPT record to an existing
-- // IAMS product record whose BVIAppStatus is not "Cancelled."

-- // Rule for product link
-- // First, determine whether to base link on Notation or ActivityID.
-- // Then, determine whether the possibly linked record has a BVIAppStatus that is not "Cancelled"


--/*
CREATE   FUNCTION ufn_GetBVIAppStatus
(
	@ActivityID UniqueIdentifier,
	@Notation varchar(1000),
	@AppID varchar(30)
)

RETURNS varchar(30)

AS
--*/

/*
	Declare @ActivityID UniqueIdentifier
	Declare @Notation varchar(1000)
	Declare @AppID varchar(30)
	Set @Input = '3981bc36-c5f5-458a-a8ef-5afffb0003ee'
*/


BEGIN

	Declare @PassIAMSLinkTestInd bit
	Declare @Notation2 varchar(32)
	Declare @BVIAppStatus varchar(30)
	Declare @GuidSoFar bit
	Declare @ValidRecordCount int
	Declare @i int

	-- // STEP 1: DETERMINE WHETHER NOTATION IS A GUID

	Set @GuidSoFar = 1
	Set @PassIAMSLinkTestInd = 0

	-- // Test length
	If (@Notation IS NULL) OR  (Len(@Notation) < 36)
	Begin
		Set @GuidSoFar = 0
	End



	-- // Test hyphens
	If @GuidSoFar = 1
	Begin

		-- // Truncate @Notation
		Set @Notation  = Left(@Notation, 36)

		If CharIndex(Substring(@Notation, 9, 1), '-') <> 1  Set @GuidSoFar = 0

		If @GuidSoFar = 1
		Begin
			If CharIndex(Substring(@Notation, 14, 1), '-') <> 1  Set @GuidSoFar = 0
		End

		If @GuidSoFar = 1
		Begin
			If CharIndex(Substring(@Notation, 19, 1), '-') <> 1  Set @GuidSoFar = 0
		End

		If @GuidSoFar = 1
		Begin
			If CharIndex(Substring(@Notation, 24, 1), '-') <> 1  Set @GuidSoFar = 0
		End
	End


	-- // Test hex
	If @GuidSoFar = 1
	Begin
		Set @Notation2 = Replace(@Notation, '-', '')
		Set @i = 1
		While @i <= 32
		Begin
			If CharIndex(Substring(@Notation2, @i, 1), '0123456789abcdef') = 0
			Begin
				Set @GuidSoFar = 0
				Break;
			End
			Set @i = @i + 1
		End
	End


	-- // NOW THAT WE HAVE DETERMINED WHETHER TO TEST THE LINK AGAINST ACTIVITYID OR NOTATION, PERFORM THE LINK TEST.
	If @GuidSoFar = 1
	Begin
		SELECT @ValidRecordCount = Count (*) FROM IAMS..AppsAndPolsSummary WHERE ActivityID = Cast(@Notation as UniqueIdentifier) AND AppID = @AppID AND BVIAppStatus <> 'CANCELLED'
	End
	Else
	Begin
		SELECT @ValidRecordCount = Count (*) FROM IAMS..AppsAndPolsSummary WHERE ActivityID = @ActivityID AND AppID = @AppID AND BVIAppStatus <> 'CANCELLED'
	End

	-- // SET RETURN VALUE TO PASS TEST IF VALID RECORD IS FOUND
	If @ValidRecordCount > 0 
	Begin
		Set @PassIAMSLinkTestInd = 1
	End

	Return @BVIAppStatus
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE   FUNCTION dbo.ufn_GetClientJoin
(
@ClientID varchar(20),
@Prefix varchar(10)
)
RETURNS varchar(60)
AS
BEGIN
	DECLARE @Results varchar(60)

	IF @ClientID IN ('MORGANS', 'HARDROCK')
		SELECT @Results = ' INNER JOIN Morgans..Employee e ON ' + @Prefix + '.EmpID = e.EmpID '
	ELSE
		SELECT @Results = ''

	RETURN @Results
END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE      FUNCTION dbo.ufn_GetClientWhere
(
@ClientID varchar(20),
@Prefix varchar(10)
)
RETURNS varchar(65)
AS
BEGIN
	DECLARE @Results varchar(65)

	IF @ClientID = 'MORGANS'
		SELECT @Results = @Prefix + '.ClientID = ''MORGANS'' AND e.GroupCode <> ''XTA'' AND '
	ELSE IF @ClientID = 'HARDROCK'
		SELECT @Results = @Prefix + '.ClientID = ''MORGANS'' AND e.GroupCode = ''XTA'' AND '
	ELSE
		SELECT @Results = @Prefix + '.ClientID = ''' + @ClientID + ''' AND '

	RETURN @Results
END








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE        FUNCTION dbo.ufn_GetClientWhereSelect
(
@ClientID varchar(20),
@Prefix varchar(10),
@HoursSection bit
)
RETURNS varchar(65)
AS
BEGIN
	DECLARE @Results varchar(65)

	IF @ClientID = 'MORGANS'
	Begin
		If @HoursSection = 1
		Begin
			SELECT @Results = @Prefix + '.ClientID = ''MORGANS'' AND '
		End
		Else
		Begin
			SELECT @Results = @Prefix + '.ClientID = ''MORGANS'' AND e.GroupCode <> ''XTA'' AND '
		End
	End
	ELSE IF @ClientID = 'HARDROCK'
	Begin
		If @HoursSection = 1
		Begin
			SELECT @Results = @Prefix + '.ClientID = ''MORGANS'' AND '
		End
		Else
		Begin
			SELECT @Results = @Prefix + '.ClientID = ''MORGANS'' AND e.GroupCode = ''XTA'' AND '
		End
	End
	ELSE IF  CharIndex('|', @ClientID) = 0
	Begin
		SELECT @Results = @Prefix + '.ClientID = ''' + @ClientID + ''' AND '
	End
	ELSE IF  CharIndex('|', @ClientID) > 0
	Begin

		Declare @SubClientID varchar(20)
		Declare @ClientRecordcount int
		Declare @ClientIterator int
		Declare @SplitOn varchar(1)
		Declare @ClientList varchar(200)
		Declare @ClientTable TABLE(RecID int, SubClientID varchar(20))
		Declare @ClientRecID int
		Declare @ColumnName varchar(20)

		Set @SplitOn = '|'
		Set @ClientRecID = 1
		Set @ClientList = @ClientID

		While (CharIndex(@SplitOn, @ClientList) > 0)
		Begin
			Insert Into @ClientTable (RecID, SubClientID)  Select RecID = @ClientRecID, SubClientID = ltrim(rtrim(Substring(@ClientList, 1, CharIndex(@SplitOn, @ClientList) -1) ) )
			Set @ClientList = Substring(@ClientList, Charindex(@SplitOn, @ClientList)+1, len(@ClientList))
			Select @ClientRecID = @ClientRecID + 1
		End
		INSERT INTO @ClientTable (RecID, SubClientID) Select RecID = @ClientRecID, SubClientID = ltrim(rtrim(@ClientList))

		Set @ClientIterator = 1
		Select @ClientRecordcount = Count (*) FROM @ClientTable	

		WHILE @ClientIterator <= @ClientRecordcount
		Begin
			SELECT @SubClientID = SubClientID FROM @ClientTable WHERE RecID = @ClientIterator
			If @ClientIterator = 1
			Begin
				Select @Results = '(' + @Prefix + '.ClientID = ''' + @SubClientID + ''' OR '
			End
			Else If @ClientIterator < @ClientRecordCount
			Begin
				Select @Results = @Results  + @Prefix + '.ClientID = ''' + @SubClientID + ''' OR '
			End
			Else If @ClientIterator = @ClientRecordCount
			Begin
				Select @Results = @Results  + @Prefix + '.ClientID = ''' + @SubClientID + ''') AND '
			End
			Set @ClientIterator = @ClientIterator + 1
		End

	End

	RETURN @Results
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   FUNCTION ufn_GetDaysInMonth
(
	@Year varchar(4),
	@Month int
)

RETURNS int

AS

BEGIN
    Return case 
	when @Month IN (1, 3, 5, 7, 8, 10, 12) then 31
             when @Month IN (4, 6, 9, 11) then 30
             else  
		case when @Year % 4  = 0 AND (@Year % 100 != 0 OR @Year % 400  = 0) then 29
                           else 28
                           end
          	end
END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





--Select Test = dbo.ufn_GetEnrolled('gstuckey', 'Martinrea', '11/17/2009')


CREATE    FUNCTION ufn_GetEnrolled
(
	@UserID varchar(20),
	@ClientID varchar(20),
	@CurDate datetime
)

RETURNS int

AS

BEGIN

	Declare @Results int

	Select @Results  = 
	(SELECT Count (*)  
	FROM EmpTransmittal etPrimary 
	INNER JOIN 
	(
	SELECT DISTINCT ept2.ActivityID 
	FROM EmpProductTransmittal ept2 
	INNER JOIN ProjectReports..EmpTransmittal et2 ON ept2.ActivityID = et2.ActivityID 
	LEFT JOIN ClientProduct_Extended cpe on ept2.ClientId = cpe.ClientId AND ept2.ProductId = cpe.ClientProductID 
	WHERE ept2.LogicalDelete = 0 AND  et2.ClientID = @ClientID AND cpe.ExtendedInd = 1 AND
	dbo.ufn_IsDateBetween(@CurDate, cpe.StartDate, cpe.EndDate) = 1 AND
	ept2.LicensedEnroller = @UserID AND dbo.ufn_IsDateEqual(et2.CallStartTime, @CurDate) = 1 AND
	et2.SupervisorApprovalDate IS NOT NULL 
	
	UNION
	
	SELECT DISTINCT v.etActivityID
	FROM v_PR_IAMS v
	WHERE v.eptLogicalDelete = 0 AND v.etClientID = @ClientID AND v.apsBVIAppStatus <> 'CANCELLED' AND
	v.eptLicensedEnroller = @UserID AND dbo.ufn_IsDateEqual(v.etCallStartTime, @CurDate) = 1 AND
	v.etSupervisorApprovalDate IS NOT NULL

	) 
	eptSecondary on etPrimary.ActivityID = eptSecondary.ActivityID
	)

	Return @Results
END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/*
   Select Results = dbo.ufn_GetPreferredAPSRelationCode('pr', EE')
*/


CREATE   FUNCTION ufn_GetPreferredAPSRelationCode
(
	@InputRelationCode varchar(10),
	@ConvertFromPR bit,
	@ConvertFromAPS bit
)


RETURNS varchar(10)

AS

BEGIN
	Declare @Results varchar(10)

	If @ConvertFromAPS = 1
	Begin
		SELECT @Results = rcm2.APS_RelationCode
		FROM RelationCodesMapping rcm
		INNER JOIN RelationCodesMapping rcm2 ON rcm.PR_RelationCode = rcm2.PR_RelationCode AND rcm2.APS_PreferredInd = 1 
		WHERE rcm.APS_RelationCode = @InputRelationCode
	End

	If @ConvertFromPR = 1
	Begin
		SELECT @Results = APS_RelationCode 
		FROM RelationCodesMapping WHERE PR_RelationCode = @InputRelationCode AND APS_PreferredInd = 1
	End

	Return @Results
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE FUNCTION ufn_GetProductFieldsTable (@ProductID varchar(20))
	RETURNS @ProductFieldsTable TABLE (RecID int, FieldName varchar(20))  
AS

BEGIN
	Declare @RecID int
	Declare @SplitOn varchar(1)
	Declare @FieldList varchar(200)

	Select @FieldList = Columns FROM Excel_SegmentConfigure WHERE SegmentId = @ProductID
	Set @SplitOn = '|'
	Set @RecID = 1
	While (CharIndex(@SplitOn, @FieldList) > 0)
	Begin
		Insert Into @ProductFieldsTable (RecID, FieldName)  Select RecID = @RecID, FieldName = ltrim(rtrim(Substring(@FieldList, 1, CharIndex(@SplitOn, @FieldList) -1) ) )
		Set @FieldList = Substring(@FieldList, Charindex(@SplitOn, @FieldList)+1, len(@FieldList))
		Set @RecID = @RecID + 1
	End
	INSERT INTO @ProductFieldsTable (RecID, FieldName) Select RecID = @RecID, FieldName = ltrim(rtrim(@FieldList))
	RETURN
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE      FUNCTION dbo.ufn_GetSqlBoilerplate
(
@ClientID varchar(20),
@Prefix varchar(10),
@APDInd bit
)
RETURNS varchar(1200)
AS
BEGIN
	DECLARE @Results varchar(1200)

	SELECT @Results = ' (SELECT Count (*) FROM IAMS..AppsAndPolsSummary aps '
	SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpTransmittal et2 on aps.ActivityID = et2.ActivityID '

	-- // 1/25/2011
	--SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpProductTransmittal ept on aps.ActivityID = ept.ActivityID AND aps.AppID = ept.AppID '
	SELECT @Results = @Results + 'INNER JOIN ProjectReports..EmpProductTransmittal ept on aps.AppID = ept.AppID '

	IF @APDInd = 1
		SELECT @Results = @Results + dbo.ufn_GetAppsAndPolsDataJoin()

	SELECT @Results = @Results + 'INNER JOIN BVI..Client bviclient ON aps.ClientID = bviclient.ClientID '
	SELECT @Results = @Results + dbo.ufn_GetClientJoin(@ClientID, 'et2')
	SELECT @Results = @Results + 'WHERE ept.LogicalDelete = 0 AND et2.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Results = @Results +  dbo.ufn_GetClientWhere(@ClientID, 'et2')
	SELECT @Results = @Results + ' aps.LicensedEnroller = u.UserID AND aps.BVIAppStatus <> ''CANCELLED'' AND '
return @results

END








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE  FUNCTION ufn_GetTableFromList (@SegmentID varchar(200))
	RETURNS @FieldsTable TABLE (RecID int, FieldName varchar(20))  
AS

BEGIN
	Declare @RecID int
	Declare @SplitOn varchar(1)
	Declare @FieldList varchar(200)

	Select @FieldList = Columns FROM Excel_SegmentConfigure WHERE SegmentId = @SegmentID
	Set @SplitOn = '|'
	Set @RecID = 1
	While (CharIndex(@SplitOn, @FieldList) > 0)
	Begin
		Insert Into @FieldsTable (RecID, FieldName)  Select RecID = @RecID, FieldName = ltrim(rtrim(Substring(@FieldList, 1, CharIndex(@SplitOn, @FieldList) -1) ) )
		Set @FieldList = Substring(@FieldList, Charindex(@SplitOn, @FieldList)+1, len(@FieldList))
		Set @RecID = @RecID + 1
	End
	INSERT INTO @FieldsTable (RecID, FieldName) Select RecID = @RecID, FieldName = ltrim(rtrim(@FieldList))
	RETURN
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE          FUNCTION dbo.ufn_GetTier
(
	@AppID varchar(30)
)
RETURNS varchar(50)
AS
BEGIN
	DECLARE @Tier varchar(50)
	SET @Tier = ''

	SELECT @Tier = case
	when (aps.ProductID = 'ALLSTATEUL') AND (apd.FieldName = 'Future Purchase Option Rider') then 'EZ'
	when (aps.ProductID = 'TMARKCCI')  AND (CharIndex('EZV', apd.FieldData) > 0) then 'EZ1_5'
	when (aps.ProductID = 'TMARKUL')  AND (CharIndex('EZV1', apd.FieldData) > 0) then 'EZ1_5'
	when (aps.ProductID = 'TMARKUL')  AND (CharIndex('EZV2', apd.FieldData) > 0) then 'EZ1_10'
	when (aps.ProductID = 'TMARKUL')  AND (CharIndex('EZV3', apd.FieldData) > 0) then 'EZ2_5'
	--when (aps.ProductID = 'TMARKCOMBO')  AND (CharIndex('EZV', apd.FieldData) > 0) then 'EZ1_5'
	when (aps.ProductID = 'TMARKCOMBO')  AND (apd.FieldName = 'EZValue' AND apd.FieldData = 'true') then 'EZ1_5'
	when (aps.ProductID = 'TRANSUL')  AND (apd.FieldName = 'EZValue')  then 'EZ'
	when (aps.ProductID = 'STANTEC')  AND (apd.FieldName = 'EZValue')  then 'EZ'
	else ''
	end

	FROM IAMS..AppsAndPolsSummary aps 
	LEFT JOIN IAMS..AppsAndPolsData apd ON aps.AppID = apd.AppID
	AND
	(
	(aps.ProductID IN ('TRANSUL', 'TMARKCOMBO')  AND apd.FieldName='EZValue' and apd.FieldData = 'true')
	OR
	(aps.ProductID = 'TMARKUL'  AND CharIndex('EZV', apd.FieldData) > 0)
	OR
	(aps.ProductID = 'ALLSTATEUL' AND apd.FieldName = 'Future Purchase Option Rider')
	) 		
	WHERE 
	aps.AppID = @AppID

	Return @Tier
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/*
   Select Results = dbo.ufn_IsAPSLegal('3981bc36-c5f5-458a-a8ef-5afffb0003ee', '22', 'Martinrea', 'TMARKCOMBO)
*/


CREATE     FUNCTION ufn_IsAPSLegal
(
	@ActivityID UniqueIdentifier,
	@AppID varchar(30),
	@ClientID varchar(20),
	@ClientProductID varchar(20)
)

RETURNS bit

AS

BEGIN
	Declare @IsAPSLegalInd bit
	Declare @IsStandardInd bit
	Declare @IsExtendedInd bit
	Declare @IsAcesSpecialInd bit

	-- // DETERMINE PRODUCT TYPE
	If (SELECT Count (*) FROM ClientProduct_Extended WHERE ClientID = @ClientID AND ClientProductID = @ClientProductID AND ExtendedInd = 1) > 0
		Set @IsExtendedInd = 1
	Else If  (SELECT Count (*) FROM ClientProduct_Extended WHERE ClientID = @ClientID AND ClientProductID = @ClientProductID AND AcesSpecialInd = 1) > 0
		Set @IsAcesSpecialInd = 1
	Else
		Set @IsStandardInd = 1

	-- // CHECK PRODUCT TYPE AGAINST APS
	If @IsExtendedInd = 1
		Set @IsAPSLegalInd = 1
	Else If @IsStandardInd = 1 OR @IsAcesSpecialInd = 1
	Begin

		-- // 1/25/2011
		--If (SELECT Count (*) FROM IAMS..AppsAndPolsSummary WHERE ActivityID = @ActivityID AND AppID = @AppID AND BVIAppStatus <> 'CANCELLED') > 0
		If (SELECT Count (*) FROM IAMS..AppsAndPolsSummary WHERE AppID = @AppID AND BVIAppStatus <> 'CANCELLED') > 0
			Set @IsAPSLegalInd = 1
		Else
			Set @IsAPSLegalInd = 0
	End
		
	Return @IsAPSLegalInd
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



--select dog = dbo.ufn_IsDateBetween('2/5/2011', '2/1/2011', '2/28/2011')

CREATE     FUNCTION ufn_IsDateBetween
(
	@DateBeingTested datetime,
	@StartDate datetime,
	@EndDate datetime
)

RETURNS int

AS



BEGIN

/*	
	DECLARE @DateBeingTested datetime
	DECLARE @StartDate datetime
	DECLARE @EndDate datetime
	SELECT @DateBeingTested = Cast('10/01/2009' as datetime)
	SELECT @StartDate  = Cast('10/7/2009' as datetime)
	SELECT @EndDate  = Cast('10/31/2009' as datetime)
*/

	DECLARE @Results int
	DECLARE @DateBeingTestedVar varchar(8)
	DECLARE	@StartDateVar varchar(8)
	DECLARE	@EndDateVar varchar(8)

	SELECT @DateBeingTestedVar = 
	Cast(Datepart(YEAR, @DateBeingTested) as varchar(4)) +
	Replace(STR(Cast(Datepart(month, @DateBeingTested)as varchar(2)), 2), ' ', '0') +
	Replace(STR(Cast(Datepart(day, @DateBeingTested)as varchar(2)), 2), ' ', '0')

	SELECT @StartDateVar = 
	Cast(Datepart(YEAR, @StartDate) as varchar(4)) +
	Replace(STR(Cast(Datepart(month, @StartDate)as varchar(2)), 2), ' ', '0') +
	Replace(STR(Cast(Datepart(day, @StartDate)as varchar(2)), 2), ' ', '0')

	SELECT @EndDateVar = 
	Cast(Datepart(YEAR, @EndDate) as varchar(4)) +
	Replace(STR(Cast(Datepart(month, @EndDate)as varchar(2)), 2), ' ', '0') +
	Replace(STR(Cast(Datepart(day, @EndDate)as varchar(2)), 2), ' ', '0')



	IF @EndDate IS NOT NULL
	Begin
		IF (@DateBeingTestedVar >= @StartDateVar) and (@DateBeingTestedVar <= @EndDateVar)
		Begin
			SELECT @Results = 1
		End
		Else	
		Begin
			SELECT @Results = 0
		End 
	End
	Else
	Begin
		IF @DateBeingTestedVar >= @StartDateVar
		Begin
			SELECT @Results = 1
		End
		Else	
		Begin
			SELECT @Results = 0
		End 
	End








	Return @Results

	--Select Results = @Results
	--select DateBeingTested = @DateBeingTestedvar
	--select StartDate = @StartDateVar
	--Select EndDate = @EndDateVar
	-- Convert 14 to a string, length = 5. Then replace blanks with 0.
	--select Value =  REPLACE(STR(14, 5), ' ', '0')

END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



-- select results =dbo.ufn_IsDateEqual('3/17/2010', '3/18/2010')

CREATE   FUNCTION ufn_IsDateEqual
(
	@Date1 datetime,
	@Date2 datetime
)

RETURNS int

AS

BEGIN

	Declare @Results int
 	If (DatePart(year, @Date1) =   DatePart(year, @Date2)) AND (DatePart(month, @Date1) = DatePart(month, @Date2)) AND (DatePart(day, @Date1) =  DatePart(day, @Date2))
	Begin
		Select @Results = 1
	End
	Else
	Begin
		Select @Results = 0
	End 
	Return @Results
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


--   Select Results = dbo.ufn_IsGuid('3981bc36-c5f5-458a-a8ef-5afffb0003ee')
--   Select Results = dbo.ufn_IsGuid('')
--   Select Results = dbo.ufn_IsGuid(null)


--/*
CREATE  FUNCTION ufn_IsGuid
(
	@Input varchar(200)
)

RETURNS bit

AS
--*/

/*
	Declare @Input varchar(200)
	Set @Input = '3981bc36-c5f5-458a-a8ef-5afffb0003ee'
*/


BEGIN

	Declare @Results bit
	Declare @i int


	-- // Test length
	If (@Input IS NULL) OR  (Len(@Input) <> 36)
	Begin
		Set @Results = 0
		Return @Results
	End


	-- // Test hyphens
	Set @Results = 1
	If CharIndex(Substring(@Input, 9, 1), '-') <> 1  Set @Results = 0
	If CharIndex(Substring(@Input, 14, 1), '-') <> 1  Set @Results = 0
	If CharIndex(Substring(@Input, 19, 1), '-') <> 1  Set @Results = 0
	If CharIndex(Substring(@Input, 24, 1), '-') <> 1  Set @Results = 0
	If @Results = 0  Return @Results

	-- // Test hex
	Set @Input = Replace(@Input, '-', '')
	Set @i = 1
	While @i <= Len(@Input)
	Begin
		If CharIndex(Substring(@Input, @i, 1), '0123456789abcdef') = 0
		Begin
			Set @Results = 0
			Return @Results
		End
		Set @i = @i + 1
	End

	Return @Results
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO













CREATE           FUNCTION ufn_IsTestID (@ClientID varchar(20), @EmpID varchar(20))

RETURNS Int AS

BEGIN

/*
Declare @ClientID varchar(20)
Declare @EmpID varchar(20)
Set @ClientID = 'BureauVeritas'
Set @EmpID = 'AS72020511'
*/

Declare @Results int
Select @Results = 0

 IF  @ClientID = 'BureauVeritas' AND (SELECT Count (*) FROM BureauVeritas..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'CTCA' AND (SELECT Count (*) FROM CTCA..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'Fulton' AND (SELECT Count (*) FROM Fulton..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'Genesis' AND (SELECT Count (*) FROM TestID_Inactive WHERE ClientID = 'Genesis' AND EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'HT' AND (SELECT Count (*) FROM TestID_Inactive WHERE ClientID = 'HT' AND EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'HARRISTEETER' AND (SELECT Count (*) FROM TestID_Inactive WHERE ClientID = 'HT' AND EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'Martinrea' AND (SELECT Count (*) FROM Martinrea..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'Morgans' AND (SELECT Count (*) FROM Morgans..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'PKOH' AND (SELECT Count (*) FROM PKOH..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'RDS' AND (SELECT Count (*) FROM RDS..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'SFWMD' AND (SELECT Count (*) FROM SFWMD..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'Stantec' AND (SELECT Count (*) FROM Stantec..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'WeatherShield' AND (SELECT Count (*) FROM WeatherShield..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'IPBC' AND (SELECT Count (*) FROM IPBC..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'DIOPITT' AND (SELECT Count (*) FROM DIOPITT..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'C3' AND (SELECT Count (*) FROM C3..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1
ELSE IF  @ClientID = 'GNC' AND (SELECT Count (*) FROM GNC..TestIDs WHERE EmpID = @EmpID) = 1
	Select @Results = 1


RETURN @Results
--Select Results = @Results
END












GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE FUNCTION ufn_LocalToEST
(
	@LocalTime datetime,
	@LocalGreenwichOffset int
)

RETURNS datetime

AS


BEGIN


/*
	DECLARE @LocalTime datetime
	DECLARE @LocalGreenwichOffset int
	SELECT @LocalTime  = '6/25/2009 12:06:00 PM'
	SELECT @LocalGreenwichOffset  = 7
*/

	DECLARE @Results datetime
	DECLARE @TotalOffset int
	DECLARE @GreenwichESTOffset int
	SELECT @GreenwichESTOffset = -5
	SET @TotalOffset = @LocalGreenwichOffset + @GreenwichESTOffset


	IF @LocalTime IS NULL 
	Begin
		SELECT @Results = NULL
	End
	Else
	Begin
		Select @Results = DATEADD(hour, @TotalOffset, @LocalTime)
	End
	Return @Results
	--Select Results = @Results


END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION dbo.ufn_PadLeft 
(
 @InputString varchar(250),
 @Length int,
 @PadChar varchar(1)
)
RETURNS varchar(1000)
AS
BEGIN
	DECLARE @Result varchar(1000)
	SELECT @Result = REPLICATE(@PadChar, @Length - LEN(ISNULL(@InputString,''))) + ISNULL(@InputString, '')
	RETURN @Result
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION dbo.ufn_PadRight
(
 @InputString varchar(250),
 @Length int,
 @PadChar varchar(1)
)
RETURNS varchar(1000)
AS
BEGIN
	DECLARE @Result varchar(1000)
  	SELECT @Result = ISNULL(@InputString, '') + REPLICATE(@PadChar, @Length -  LEN(ISNULL(@InputString,'')))
	RETURN @Result
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



--   Select Results = dbo.ufn_IsGuid('3981bc36-c5f5-458a-a8ef-5afffb0003ee')
--   Select Results = dbo.ufn_IsGuid('')
--   Select Results = dbo.ufn_IsGuid(null)

--   Select Results = dbo.ufn_GetBVIAppStatus('7322bec5-06c3-44f3-a50a-cd5b05fc7506', '7322bec5-06c3-44f3-a50a-cd5b05fc7506', '20091118134835310')


-- // Rule for call
-- // Call is legal whether or not IAMS call record exists.
-- // Convert-to-call calls have no bearing on this rule.

-- // Rule for product
-- // Product is legal only when it is possible to link the EPT record to an existing
-- // IAMS product record whose BVIAppStatus is not "Cancelled."

-- // Rule for product link
-- // First, determine whether to base link on Notation or ActivityID.
-- // Then, determine whether the possibly linked record has a BVIAppStatus that is not "Cancelled"


--/*
CREATE   FUNCTION ufn_PassIAMSLinkTest
(
	@ActivityID UniqueIdentifier,
	@Notation varchar(1000),
	@AppID varchar(30)
)

RETURNS varchar(30)

AS
--*/

/*
	Declare @ActivityID UniqueIdentifier
	Declare @Notation varchar(1000)
	Declare @AppID varchar(30)
	Set @Input = '3981bc36-c5f5-458a-a8ef-5afffb0003ee'
*/


BEGIN

	Declare @PassIAMSLinkTestInd bit
	Declare @Notation2 varchar(32)
	Declare @BVIAppStatus varchar(30)
	Declare @GuidSoFar bit
	Declare @ValidRecordCount int
	Declare @i int

	-- // STEP 1: DETERMINE WHETHER NOTATION IS A GUID

	Set @GuidSoFar = 1
	Set @PassIAMSLinkTestInd = 0

	-- // Test length
	If (@Notation IS NULL) OR  (Len(@Notation) < 36)
	Begin
		Set @GuidSoFar = 0
	End



	-- // Test hyphens
	If @GuidSoFar = 1
	Begin

		-- // Truncate @Notation
		Set @Notation  = Left(@Notation, 36)

		If CharIndex(Substring(@Notation, 9, 1), '-') <> 1  Set @GuidSoFar = 0

		If @GuidSoFar = 1
		Begin
			If CharIndex(Substring(@Notation, 14, 1), '-') <> 1  Set @GuidSoFar = 0
		End

		If @GuidSoFar = 1
		Begin
			If CharIndex(Substring(@Notation, 19, 1), '-') <> 1  Set @GuidSoFar = 0
		End

		If @GuidSoFar = 1
		Begin
			If CharIndex(Substring(@Notation, 24, 1), '-') <> 1  Set @GuidSoFar = 0
		End
	End


	-- // Test hex
	If @GuidSoFar = 1
	Begin
		Set @Notation2 = Replace(@Notation, '-', '')
		Set @i = 1
		While @i <= 32
		Begin
			If CharIndex(Substring(@Notation2, @i, 1), '0123456789abcdef') = 0
			Begin
				Set @GuidSoFar = 0
				Break;
			End
			Set @i = @i + 1
		End
	End


	-- // NOW THAT WE HAVE DETERMINED WHETHER TO TEST THE LINK AGAINST ACTIVITYID OR NOTATION, PERFORM THE LINK TEST.
	If @GuidSoFar = 1
	Begin
		SELECT @ValidRecordCount = Count (*) FROM IAMS..AppsAndPolsSummary WHERE ActivityID = Cast(@Notation as UniqueIdentifier) AND AppID = @AppID AND BVIAppStatus <> 'CANCELLED'
	End
	Else
	Begin
		SELECT @ValidRecordCount = Count (*) FROM IAMS..AppsAndPolsSummary WHERE ActivityID = @ActivityID AND AppID = @AppID AND BVIAppStatus <> 'CANCELLED'
	End

	-- // SET RETURN VALUE TO PASS TEST IF VALID RECORD IS FOUND
	If @ValidRecordCount > 0 
	Begin
		Set @PassIAMSLinkTestInd = 1
	End

	Return @BVIAppStatus
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO









CREATE       FUNCTION dbo.ufn_RptBuildProductSegment3
(
@SourceInd bit,
@ClientID varchar(20),
@ProductID varchar(20),
@Selection varchar(200),
@TestedDate varchar(20),
@CallDate smalldatetime,
@AddlCondition varchar(400),
@APDJoinInd bit,
@Prefix varchar(10)
)
RETURNS varchar(4000)
AS
BEGIN

	DECLARE @Source varchar(9)
	DECLARE @Results varchar(4000)

	If @SourceInd = 0
		Set @Source = 'standard'
	Else	
		Set @Source = 'extended'



	SELECT @Results = '(SELECT ' + @Selection + ' ' 

	If @Source = 'standard'
	Begin
		SELECT @Results = @Results + 'FROM  EmpProductTransmittal ept '
	End
	Else
	Begin
		SELECT @Results = @Results + 'FROM  Alt_ProductData alt '
		SELECT @Results = @Results + 'INNER JOIN EmpProductTransmittal ept on alt.AltProductDataID = ept.AltProductDataID '
	End

	SELECT @Results = @Results + 'INNER JOIN EmpTransmittal et2 on ept.ActivityID = et2.ActivityID '
	SELECT @Results = @Results + 'INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID '
	SELECT @Results = @Results + dbo.ufn_GetClientJoin(@ClientID, @Prefix)

	-- // Standard links to IAMS
	If @Source = 'standard'
	Begin

		-- // 1/25/2011
		--SELECT @Results = @Results + 'INNER JOIN IAMS..AppsAndPolsSummary aps ON ept.ActivityID = aps.ActivityID AND ept.AppID = aps.AppID '
		SELECT @Results = @Results + 'INNER JOIN IAMS..AppsAndPolsSummary aps ON ept.AppID = aps.AppID '
		IF @APDJoinInd = 1
			SELECT @Results = @Results + dbo.ufn_GetAppsAndPolsDataJoin()
	End

	SELECT @Results = @Results + 'WHERE ept.LogicalDelete = 0 AND '
	SELECT @Results = @Results + 'et2.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Results = @Results + dbo.ufn_GetClientWhere(@ClientID, @Prefix)
	SELECT @Results = @Results + ' ept.LicensedEnroller = u.UserID AND '
	SELECT @Results = @Results + 'ept.ProductID = ''' + @ProductID + ''' AND '

	If @Source = 'standard'
		SELECT @Results = @Results + ' aps.BVIAppStatus <> ''CANCELLED'' AND '

	-- 11/27/2009
	If @Source = 'extended'
		SELECT @Results = @Results + 'dbo.ufn_IsAPSLegal(ept.ActivityID, ept.AppID, ept.ClientID, ept.ProductID) = 1 AND '


	IF @AddlCondition <> ''
		SELECT @Results = @Results + @AddlCondition + ' AND '

	SELECT @Results = @Results + 'dbo.ufn_IsDateEqual(' + Convert(varchar, @TestedDate, 101) + ', ''' + Convert(varchar, @CallDate, 101) + ''') = 1 AND '
	SELECT @Results = @Results  + 'dbo.ufn_IsTestID(bviclient.ClientID, et2.EmpID) = 0) '

	RETURN @Results
END














GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE  FUNCTION ufn_SalesByStore (@storeid varchar(30))
   	RETURNS @t TABLE (title varchar(80) NOT NULL, qty smallint NOT NULL)  
AS
BEGIN
	   INSERT INTO @t (title, qty) values ('hello there', 4)
	   RETURN
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO






CREATE  FUNCTION dbo.ufn_Split
(
	@RowData varchar(2000),
	@SplitOn varchar(1)
)  
RETURNS @RtnValue table 
(
	--Id int identity(1,1),
	Data varchar(100)
) 
AS  
BEGIN 
	Declare @Cnt int
	Set @Cnt = 1

	While (Charindex(@SplitOn,@RowData)>0)
	Begin
		Insert Into @RtnValue (data)
		Select 
			Data = ltrim(rtrim(Substring(@RowData,1,Charindex(@SplitOn,@RowData)-1)))

		Set @RowData = Substring(@RowData,Charindex(@SplitOn,@RowData)+1,len(@RowData))
		Set @Cnt = @Cnt + 1
	End
	
	Insert Into @RtnValue (data)
	Select Data = ltrim(rtrim(@RowData))

	Return
END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION ufn_StringCleanser (@Input VarChar(2048), @OutputRange VarChar(10)) 
RETURNS varchar(2048)
AS 

/*
	Strip out non-numeric characters
	SELECT dbo.ufn_StringCleanser('~`0~1@2#3$4%5^6&7*8(9)0_a-b+c=d','0-9')
	SELECT dbo.ufn_StringCleanser('~`0~1@2#3$4%5^6&7*8(9)0_a-b+c=d','A-Z')
	SELECT dbo.ufn_StringCleanser('~`0~1@2#3$4%5^6&7*8(9)0_a-b+c=d','0-9,A-Z')

*/

BEGIN 
WHILE PATINDEX('%[^'+@OutputRange+']%', @Input) > 0 
BEGIN 
SET @Input = STUFF(@Input, PATINDEX('%[^'+@OutputRange+']%', @Input), 1, '') 
END 
RETURN @Input 
END 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


--Select Results = dbo.ufn_ToInt('88.24')


   CREATE FUNCTION dbo.ufn_ToInt(@Input varchar(100))

RETURNS  int

AS  
BEGIN 
	Declare @Cnt int
	Declare @CurChar varchar(1)
	Declare @OutputVarchar varchar(100)
	Declare @OutputInt int

	Set @OutputVarchar = ''
	Set @OutputInt = 0
	Set @Cnt = 1

	While (@Cnt < len(@Input) + 1)
	Begin
		Select @CurChar = Substring(@Input, @Cnt, 1)

		If CharIndex(@CurChar, '0123456789') > 0
		Begin
			Set @OutputVarchar = @OutputVarchar + @CurChar
		End
		Else
		Begin
			break;
		End

		Set @Cnt = @Cnt + 1
	End

	If len(@OutputVarchar) > 0 
		Set @OutputInt = Convert(int, @OutputVarchar)
		

	Return @OutputInt
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  FUNCTION dbo.ufn_ToProper
(
 @Input varchar(8000)
)
RETURNS varchar(8000)
AS
BEGIN
	DECLARE @Reset bit;
	DECLARE @Ret varchar(8000);
	DECLARE @i int;
	DECLARE @c char(1);

	SET @Reset  = 1
	SET @i = 1
	SET @Ret = ''

	WHILE (@i <= len(@Input))
	Begin 
		SELECT @c= substring(@Input, @i, 1),
		@Ret = @Ret + case when @Reset=1 then UPPER(@c) else LOWER(@c) end,
		@Reset = case when @c = ' ' then 1 else 0 end,
		@i = @i +1
	End

	Return @Ret



END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO






CREATE FUNCTION [dbo].[ufn_get_FirstDayOfFollowingMonth] (@id_date datetime)  
RETURNS datetime  
AS  
BEGIN  
-- =============================================  
--  Function: dbo.ufn_get_FirstDayOfFollowingMonth  
--      - Return the first day of the next month for the input date  
--  
--  Input:    char date  
--  Output:   first of next month, or null  
-- =============================================  
--  
-- 08/22/07 - KRB: Created  
--  
--==============================================  
  
  
DECLARE @d_RetVal        datetime 
  
    set @d_RetVal = NULL   
    IF ISDATE(@id_date) = 1   
    BEGIN  
        --  get first day of following month  
        set @d_RetVal =  dateadd(m,1,CAST( CAST(month(@id_date) AS CHAR(2))  
                              + '/01/'   
                              + CAST(year(@id_date) AS CHAR(4)) AS DATETIME))
    END  
  
    RETURN(@d_RetVal)  
END  





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




Create  PROC spDbDefrag @Defrag		TinyInt
as

-- Perform a 'USE <database name>' to select 
-- 	the database in which to run the script.
--USE ACES

-- Declare variables
SET NOCOUNT ON
--DECLARE @Defrag		TinyInt			SET @Defrag = 1
DECLARE @MaxFrag	Decimal			SET @maxfrag = 10.0
DECLARE @InTableList	VarChar (128)		SET @InTableList = '%'
DECLARE @ExTableList	VarChar (128)		SET @ExTableList = '%Merge%'
DECLARE @tablename		VarChar (128)
DECLARE @execstr		VarChar (255)
DECLARE @objectid		Int
DECLARE @indexid		Int
DECLARE @frag			Decimal

--	DROP TABLE #fraglist
-- 	Declare cursor
DECLARE tables CURSOR FOR
   SELECT distinct TABLE_NAME
   FROM INFORMATION_SCHEMA.TABLES
   WHERE TABLE_TYPE = 'BASE TABLE'
	AND TABLE_NAME NOT LIKE @ExTableList
	AND TABLE_NAME LIKE @InTableList

-- Create the table
CREATE TABLE #fraglist (
   ObjectName CHAR (255), ObjectId INT, IndexName CHAR (255), IndexId INT,
   Lvl INT, CountPages INT, CountRows INT, MinRecSize INT, MaxRecSize INT,
   AvgRecSize INT, ForRecCount INT, Extents INT, ExtentSwitches INT,
   AvgFreeBytes INT, AvgPageDensity INT, ScanDensity DECIMAL, BestCount INT,
   ActualCount INT, LogicalFrag DECIMAL, ExtentFrag DECIMAL)

-- Open the cursor
OPEN tables

-- Loop through all the tables in the database
FETCH NEXT
   FROM tables
   INTO @tablename

WHILE @@FETCH_STATUS = 0
BEGIN
-- Do the showcontig of all indexes of the table
   INSERT INTO #fraglist 
   EXEC ('DBCC SHOWCONTIG (''' + @tablename + ''') WITH FAST, TABLERESULTS, ALL_INDEXES, NO_INFOMSGS')
   FETCH NEXT
      FROM tables
      INTO @tablename
END

-- Close and deallocate the 'tables' cursor
CLOSE tables
DEALLOCATE tables

IF @Defrag = 0  BEGIN
	SELECT ObjectName, ObjectId, IndexId, LogicalFrag
	FROM #fraglist
	WHERE LogicalFrag >= @MaxFrag
		AND INDEXPROPERTY (ObjectId, IndexName, 'IndexDepth') > 0
	ORDER BY LogicalFrag DESC
  END

ELSE  BEGIN
-- Declare cursor for list of indexes to be defragged
	DECLARE indexes CURSOR FOR
		SELECT ObjectName, ObjectId, IndexId, LogicalFrag
		FROM #fraglist
		WHERE LogicalFrag >= @MaxFrag
			AND INDEXPROPERTY (ObjectId, IndexName, 'IndexDepth') > 0

-- Open the cursor
	OPEN indexes

-- loop through the indexes
	FETCH NEXT
	FROM indexes
	INTO @tablename, @objectid, @indexid, @frag

	WHILE @@FETCH_STATUS = 0  BEGIN
		PRINT 'Executing DBCC INDEXDEFRAG (0, ' + RTRIM(@tablename) + ', ' + RTRIM(@indexid) + ') - fragmentation currently '
			+ RTRIM(CONVERT(varchar(15),@frag)) + '%'

		SELECT @execstr = 'DBCC INDEXDEFRAG (0, ' + RTRIM(@objectid) + ', ' + RTRIM(@indexid) + ')'

		EXEC (@execstr)
		PRINT '==========================================================================================================='
		PRINT '==========================================================================================================='
		PRINT ''
		PRINT ''

		FETCH NEXT
		FROM indexes
		INTO @tablename, @objectid, @indexid, @frag
	  END

-- Close and deallocate the 'indexes' cursor
	CLOSE indexes
	DEALLOCATE indexes
  END

-- Delete the temporary table '#fraglist'
DROP TABLE #fraglist





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





Create  PROC spDbReIndex @Defrag		TinyInt
as

/*	exec c3.dbo.spDbReIndex 1
	DBCC DBREINDEX (BenefitPlanCosts, PK_BenefitPlanCosts, 100)
	DBCC DBREINDEX (  table_name  [ , index_name [ , fillfactor ] ] )   [ WITH NO_INFOMSGS ] 
	DBCC INDEXDEFRAG ({ database_name | database_id | 0 } , { table_name | table_id | view_name | view_id } 
	[ , { index_name | index_id } [ , { partition_number | 0 } ] ]) [ WITH NO_INFOMSGS ] 



-- Perform a 'USE <database name>' to select 
-- 	the database in which to run the script.
--USE ACES
*/
-- Declare variables
SET NOCOUNT ON
--DECLARE @Defrag		TinyInt			SET @Defrag = 1
DECLARE @MaxFrag	Decimal			SET @maxfrag = 10.0
DECLARE @InTableList	VarChar (128)		SET @InTableList = '%'
DECLARE @ExTableList	VarChar (128)		SET @ExTableList = '%Merge%'
DECLARE @tablename		VarChar (128)
DECLARE @execstr		VarChar (255)
DECLARE @objectid		Int, 		@objectname	Varchar(255)
DECLARE @indexid		Int, 		@indexname	Varchar(255)
DECLARE @frag			Decimal

--	DROP TABLE #fraglist
-- 	Declare cursor
DECLARE tables CURSOR FOR
   SELECT distinct TABLE_NAME
   FROM INFORMATION_SCHEMA.TABLES
   WHERE TABLE_TYPE = 'BASE TABLE'
	AND TABLE_NAME NOT LIKE @ExTableList
	AND TABLE_NAME LIKE @InTableList

-- Create the table
CREATE TABLE #fraglist (
   ObjectName CHAR (255), ObjectId INT, IndexName CHAR (255), IndexId INT,
   Lvl INT, CountPages INT, CountRows INT, MinRecSize INT, MaxRecSize INT,
   AvgRecSize INT, ForRecCount INT, Extents INT, ExtentSwitches INT,
   AvgFreeBytes INT, AvgPageDensity INT, ScanDensity DECIMAL, BestCount INT,
   ActualCount INT, LogicalFrag DECIMAL, ExtentFrag DECIMAL)

-- Open the cursor
OPEN tables

-- Loop through all the tables in the database
FETCH NEXT
   FROM tables
   INTO @tablename

WHILE @@FETCH_STATUS = 0
BEGIN
-- Do the showcontig of all indexes of the table
   INSERT INTO #fraglist 
   EXEC ('DBCC SHOWCONTIG (''' + @tablename + ''') WITH FAST, TABLERESULTS, ALL_INDEXES, NO_INFOMSGS')
   FETCH NEXT
      FROM tables
      INTO @tablename
END

-- Close and deallocate the 'tables' cursor
CLOSE tables
DEALLOCATE tables

IF @Defrag = 0  BEGIN
	SELECT ObjectName, ObjectId, IndexName, IndexId,  LogicalFrag
	FROM #fraglist
	WHERE LogicalFrag >= @MaxFrag
		AND INDEXPROPERTY (ObjectId, IndexName, 'IndexDepth') > 0
	ORDER BY LogicalFrag DESC
  END

ELSE  BEGIN
-- Declare cursor for list of indexes to be defragged
	DECLARE indexes CURSOR FOR
		SELECT ObjectName, ObjectId, IndexName, IndexId, LogicalFrag
		FROM #fraglist
		WHERE LogicalFrag >= @MaxFrag
			AND INDEXPROPERTY (ObjectId, IndexName, 'IndexDepth') > 0

-- Open the cursor
	OPEN indexes

-- loop through the indexes
	FETCH NEXT
	FROM indexes
	INTO @tablename, @objectid, @IndexName, @indexid, @frag

	WHILE @@FETCH_STATUS = 0  BEGIN
		PRINT 'Executing DBCC DBREINDEX (' + RTRIM(@tablename) + ', ' + RTRIM(@IndexName) + ', 100) - fragmentation currently '
			+ RTRIM(CONVERT(varchar(15),@frag)) + '%'

		SELECT @execstr = 'DBCC DBREINDEX (' + RTRIM(@tablename) + ', ' + RTRIM(@IndexName) + ', 100)'

		EXEC (@execstr)
--		PRINT (@execstr)
		PRINT '==========================================================================================================='
		PRINT '==========================================================================================================='
		PRINT ''
		PRINT ''

/*
	DBCC DBREINDEX (  table_name  [ , index_name [ , fillfactor ] ] )   [ WITH NO_INFOMSGS ] 
	DBCC INDEXDEFRAG ({ database_name | database_id | 0 } , { table_name | table_id | view_name | view_id } 
	[ , { index_name | index_id } [ , { partition_number | 0 } ] ]) [ WITH NO_INFOMSGS ] 
*/

		FETCH NEXT
		FROM indexes
--		INTO @tablename, @objectid, @indexid, @frag
		Into @tablename, @objectid, @IndexName, @indexid, @frag
	  END

-- Close and deallocate the 'indexes' cursor
	CLOSE indexes
	DEALLOCATE indexes
  END

-- Delete the temporary table '#fraglist'
DROP TABLE #fraglist






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




Create  procedure spRunReplication
As

	Declare @JobName	Varchar(100), 	@ServerName Varchar(50), 	@DbName Varchar(50), @Sql Varchar(500)
	Set @ServerName = (select @@ServerName)
	SET @DbName = (select db_name())
	Set  @JobName	= (SELECT  name FROM  [hbg-sql].Distribution1.dbo.MSmerge_agents WHERE     publisher_db = @DbName)

--		Print @JobName + ' ServerName' + @ServerName + ' DbName' + @DbName
		If @ServerName LIKE '%SQL'
		BEGIN 
			exec [HBG-SQL].msdb.dbo.sp_start_job @JobName 
		End









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







--EXEC spWhere_Am_I 'PendING%D'


CREATE     procedure spWhere_Am_I
--Create   procedure spWhere_Am_I
@cString varchar(1000)
AS

/*--------------------------------------------------------------------------------------------------------


	Purpose:-If you have written stored procedures in the past where you could remember some of the code 
	they contained, but not the name of the stored procedure. 
	This tool will help you find those elusive stored procedures based on the code in the stored procedure. 
	It searches the sysobjects and syscomments tables to find any occurence of a text string. 
	It sorts the return set by object name. (Supports SQL 2000, SQL 7.0, and possibly SQL 6.5) 
--------------------------------------------------------------------------------------------------------------
*/
Declare	@cString_sql varchar(1000)
	set nocount on
	Select @cString_sql = 'select substring( o.name, 1, 50 ) as Object, count(*) as Occurences, ' +
		'case ' +
		' when o.xtype = ''D'' then ''Default'' ' +
		' when o.xtype = ''F'' then ''Foreign Key'' ' +
		' when o.xtype = ''P'' then ''Stored Procedure'' ' +
		' when o.xtype = ''PK'' then ''Primary Key'' ' +
		' when o.xtype = ''S'' then ''System Table'' ' +
		' when o.xtype = ''TR'' then ''Trigger'' ' +
		' when o.xtype = ''U'' then ''User Table'' ' +
		' when o.xtype = ''V'' then ''View'' ' +
		'end as Type ' +
		'from syscomments c join sysobjects o on c.id = o.id ' +
		'where patindex( ''%'  + @cString + '%'', c.text ) > 0 ' +
		'group by o.name, o.xtype ' +
		'order by o.xtype, o.name'

	Execute( @cString_sql )


Select @cString_sql = 'select substring( o.name, 1, 50 ) as Object, count(*) as Occurences, ' +
		'case ' +
		' when o.xtype = ''D'' then ''Default'' ' +
		' when o.xtype = ''F'' then ''Foreign Key'' ' +
		' when o.xtype = ''P'' then ''Stored Procedure'' ' +
		' when o.xtype = ''PK'' then ''Primary Key'' ' +
		' when o.xtype = ''S'' then ''System Table'' ' +
		' when o.xtype = ''TR'' then ''Trigger'' ' +
		' when o.xtype = ''U'' then ''User Table'' ' +
		' when o.xtype = ''V'' then ''View'' ' +
		'end as Type ' +
		'from syscolumns c join sysobjects o on c.id = o.id ' +
		'where patindex( ''%'  + @cString + '%'', c.name ) > 0 ' +
		'group by o.name, o.xtype ' +
		'order by o.xtype, o.name'

	Execute( @cString_sql )

Return











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



--  select * from empproducttransmittal where adddate > '3/1/2010'


/*
exec usp_CCM2010BoardRecords
*/


CREATE     PROCEDURE  usp_CCM2010BoardRecords

AS

SET NOCOUNT ON

DECLARE @ActivityID uniqueidentifier
DECLARE @AppID varchar(30)
DECLARE @AltProductDataID int
DECLARE @ProdDetailID varchar(25)
DECLARE @ClientSubProductID varchar(25)
DECLARE @RelationCode varchar(20)
DECLARE @InsuredName varchar(52)
DECLARE @Tier varchar(50)
DECLARE @FamilyProductInd bit
DECLARE @ProductID varchar(25)
DECLARE @TierRaw varchar(50)

DECLARE BobCursor CURSOR FOR Select ActivityID, AppID, AltProductDataID FROM EmpProductTransmittal

Open BobCursor

Fetch  BobCursor into @ActivityID, @AppID, @AltProductDataID

		While @@Fetch_Status = 0
		Begin
			If @AltProductDataID IS NULL
			Begin
				-- // Update from AppsAndPolsSummary
				exec ProjectReports..usp_GetConditionedFields @AppID, @ProdDetailID OUTPUT, @RelationCode OUTPUT, @InsuredName OUTPUT, @Tier OUTPUT

				UPDATE EmpProductTransmittal
				SET
				ProdDetailID = EmpProductTransmittal.ProductID,
				RelationCode = @RelationCode,
				InsuredName = @InsuredName,
				Tier = @Tier, 
				ClientSubProductID = null
				WHERE EmpProductTransmittal.AppID = @AppID
			End
			Else
			Begin
				-- // Update from Alt_ProductData
				SELECT
				@ProdDetailID = alt.ClientSubProductID,
				@ClientSubProductID = case
					when (SELECT FamilyProductInd FROM Config_ProdDetail WHERE ClientSubProductID = alt.ClientSubProductID) = 1  then alt.ClientSubProductID
					else null
					end,
				@RelationCode = alt.RelationCode,
				@InsuredName = IsNull(alt.InsuredName, ''),
				@Tier = IsNull(alt.Tier, '')
				FROM Alt_ProductData alt 
				WHERE alt.AltProductDataID = @AltProductDataID

				UPDATE EmpProductTransmittal
				SET 
				ProdDetailID = @ProdDetailID,
				ClientSubProductID = @ClientSubProductID,
				RelationCode = @RelationCode,
				InsuredName = @InsuredName,
				Tier = @Tier
				WHERE EmpProductTransmittal.AltProductDataID = @AltProductDataID	
			End
		
			Fetch  BobCursor into @ActivityID, @AppID, @AltProductDataID
		End

Close BobCursor
Deallocate BobCursor





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







/*
exec ProjectReports..usp_CCM2010BoardRecordsThisEmp 'jwatroba'

*/

CREATE Procedure usp_CCM2010BoardRecordsThisEmp
(
@LicensedEnroller varchar(25)
)

AS

BEGIN

	Declare @AppID varchar(30)
	Declare @ProdDetailID varchar(25)
	Declare @RelationCode varchar(20)
	Declare @InsuredName varchar(52)
	Declare @Tier varchar(50)

	Declare @AppTableRecID int
	Declare @AppTableRecordcount int
	Declare @AppTable Table(RecID int IDENTITY,AppID varchar(30))

	Insert Into @AppTable Select AppID FROM EmpProductTransmittal WHERE LicensedEnroller = @LicensedEnroller AND ProdDetailID IS NULL AND AddDate > '03/07/2010'
	Select @AppTableRecordcount = Count (*) FROM @AppTable

	SELECT @AppTableRecID = 1
	WHILE @AppTableRecID <= @AppTableRecordcount
	Begin	

		Select @AppID = AppID FROM @AppTable WHERE RecID = @AppTableRecID

		exec ProjectReports..usp_GetConditionedFields @AppID, @ProdDetailID OUTPUT, @RelationCode OUTPUT, @InsuredName OUTPUT, @Tier OUTPUT

		--Select @AppID, @ProdDetailID, @RelationCode, @InsuredName, @Tier


		UPDATE EmpProductTransmittal
		SET
		ProdDetailID = EmpProductTransmittal.ProductID,
		RelationCode = @RelationCode,
		InsuredName = @InsuredName,
		Tier = @Tier, 
		ClientSubProductID = null
		WHERE EmpProductTransmittal.AppID = @AppID

		Select @AppTableRecID = @AppTableRecID + 1
	End




END








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


-- //  *** TAKEN DIRECTLY FROM CALL HISTORY PAGE

CREATE  PROCEDURE usp_CallHistoryWorklistSubTableQueryAdj

AS
SELECT 
[Call Date] = et.CallStartTime,
[Emp Name] =et.LastName + ', ' + et.FirstName, 
Placeholder = '',

Enrolled = case
when ept.ProductID IS NULL then 0
else 1
end,

Product = ISNULL(ept.ProductID, ''),


Tier = case when (ept.ProductID = 'ALLSTATEUL') AND
(alt.AltProductDataID IS NULL) AND
(LEFT(apd.FieldName, 22) = 'Future Purchase Option') then 'EZ' when (ept.ProductID = 'TMARKCOMBO')  AND
(alt.AltProductDataID IS NULL) AND
(apd.FieldName='EZValue' AND
apd.FieldData = 'true') then 'EZ1_5' when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV1', apd.FieldData) > 0) then 'EZ1_5' when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV2', apd.FieldData) > 0) then 'EZ1_10' when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV3', apd.FieldData) > 0) then 'EZ2_5' when (ept.ProductID = 'TRANSUL')   AND
(alt.AltProductDataID IS NULL) AND
(apd.FieldName = 'EZValue') then 'EZ' else alt.Tier end, 

WeeklyPremium = case when ccp.DisplayPremiumInd = 1 then Cast(ept.WeeklyPremium as money) else null end,
AnnualPremium = case when ccp.DisplayPremiumInd = 1 then Cast(ept.AnnualPremium as money) else null end, 
PlanBenefitAmt = case when ccp.DisplayPlanBenefitAmtInd = 1 then Cast(ept.PlanBenefitAmt as money) else null end

 
FROM EmpTransmittal et 
LEFT JOIN EmpProductTransmittal ept on ept.ActivityID = et.ActivityID 
LEFT JOIN IAMS..AppsAndPolsSummary aps ON aps.AppID = ept.AppID AND aps.ActivityID = et.ActivityID
LEFT JOIN Config_ClientProduct ccp ON ept.ProductID = ccp.ClientProductID 

 
LEFT JOIN IAMS..AppsAndPolsData apd ON aps.AppID = apd.AppID AND ((aps.ProductID = 'ALLSTATEUL' AND
LEFT(apd.FieldName, 22) = 'Future Purchase Option') OR (aps.ProductID IN ('TRANSUL', 'TMARKCOMBO') AND
apd.FieldName='EZValue' AND apd.FieldData = 'true') OR (aps.ProductID = 'TMARKUL'  AND CharIndex('EZV', apd.FieldData) > 0) ) 

LEFT JOIN Alt_ProductData alt ON alt.AltProductDataID = ept.AltProductDataID 


WHERE 

et.ClientID = 'Diopitt'  
and dbo.ufn_IsDateBetween(et.CallStartTime, '10/27/2010', '12/13/2010') = 1


and isnull(et.LogicalDelete, 0) <> 1 and isnull(ept.LogicalDelete, 0) <> 1 and et.SupervisorApprovalDate IS NOT NULL

order by et.Lastname + et.Firstname


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE usp_DeleteCall
(
 	@ActivityID uniqueidentifier
)
AS

BEGIN TRANSACTION

	BEGIN
		-- // Delete Alt_ProductData records
		DELETE FROM Alt_ProductData WHERE ActivityID = @ActivityID
		IF @@ERROR <> 0
		BEGIN
		    	ROLLBACK
			SELECT Results = 'Error: Cannot delete Alt_ProductData data'
		    	RETURN
		END

		-- // Logically delete EmpProductTransmittal records
		--DELETE FROM EmpProductTransmittal WHERE ActivityID = @ActivityID
		UPDATE EmpProductTransmittal SET LogicalDelete = 1, AltProductDataID = null WHERE ActivityID = @ActivityID
		IF @@ERROR <> 0
		BEGIN
		    	ROLLBACK
			SELECT Results = 'Error: Cannot logically delete EmpProductTransmittal data'
		    	RETURN
		END
	
		-- // Logically delete EmpTransmittal records
		--DELETE FROM EmpTransmittal WHERE ActivityID = @ActivityID
		UPDATE EmpTransmittal SET LogicalDelete = 1 WHERE ActivityID = @ActivityID
		IF @@ERROR <> 0
		BEGIN
		    	ROLLBACK
			SELECT Results = 'Error: Cannot logically delete EmpTransmittal record'
		    	RETURN
		END
	END
		
COMMIT

SELECT Results = 'Success'





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE    PROCEDURE usp_DeleteCallMonitor
(
 	@ActivityID uniqueidentifier
)
AS

BEGIN TRANSACTION

	BEGIN
		-- // Logically delete CallMonitorChild records
		UPDATE CallMonitorChild SET LogicalDelete = 1 WHERE ActivityID = @ActivityID
		IF @@ERROR <> 0
		BEGIN
		    	ROLLBACK
			SELECT Results = 'Error: Cannot logically delete CallMonitorChild data'
		    	RETURN
		END

		-- // Logically delete CallMonitorParent record
		UPDATE CallMonitorParent SET LogicalDelete = 1 WHERE ActivityID = @ActivityID
		IF @@ERROR <> 0
		BEGIN
		    	ROLLBACK
			SELECT Results = 'Error: Cannot logically delete CallMonitorParent record'
		    	RETURN
		END
	END
		
COMMIT

SELECT Results = 'Success'






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO








/*
  exec usp_DeleteIAMSCancelled @AppID = '20111024100445093'
*/

CREATE    PROCEDURE [dbo].[usp_DeleteIAMSCancelled]
(
 @AppID varchar(30)
)
---------------------------------------------------------------
-- NAME: usp_DeleteIAMSCancelled

-- PURPOSE: Delete ProjectReports records when BVIAppStatus is cancelled

-- INPUTS:  AppID

-- OUTPUTS:  None

-- HISTORY:
--	10/25/11 - RMB - Created

----------------------------------------------------------------

AS
SET NOCOUNT ON


DECLARE @ActivityID uniqueidentifier

--  Get the ActivityID
SELECT @ActivityID = ActivityID FROM IAMS..AppsAndPolsSummary WHERE AppID = @AppID

--DELETE Alt_ProductData WHERE AppID = @AppID
--DELETE EmpProductTransmittal WHERE AppID = @AppID
--DELETE EmpTransmittal WHERE ActivityID = @ActivityID

--Select ActitivityID = @ActivityID
--select * from emptransmittal where ActivityID = @ActivityID
--select * from EmpProductTransmittal WHERE AppID = @AppID
--select * from Alt_ProductData WHERE AppID = @AppID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE usp_DeleteProduct
(
 	@ActivityID uniqueidentifier,
	@AppID varchar(30)
)
AS

BEGIN TRANSACTION

	BEGIN
		-- // Delete Alt_ProductData record
		DELETE FROM Alt_ProductData WHERE ActivityID = @ActivityID AND AppID = @AppID
		IF @@ERROR <> 0
		BEGIN
		    	ROLLBACK
			SELECT Results = 'Error: Cannot delete Alt_ProductData data'
		    	RETURN
		END

		-- // Delete EmpProductTransmittal record
		--DELETE FROM EmpProductTransmittal WHERE ActivityID = @ActivityID AND AppID = @AppID
		UPDATE EmpProductTransmittal SET LogicalDelete = 1, AltProductDataID = null WHERE ActivityID = @ActivityID AND AppID = @AppID
		IF @@ERROR <> 0
		BEGIN
		    	ROLLBACK
			SELECT Results = 'Error: Cannot logically delete EmpProductTransmittal data'
		    	RETURN
		END
	END
		
COMMIT

SELECT Results = 'Success'





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE     PROCEDURE usp_DeleteUndeleteCallMonitor
(
	@Delete bit, 
	@ActivityID uniqueidentifier
)
AS

BEGIN TRANSACTION

	BEGIN

		IF @Delete = 1
		BEGIN

			-- // Logically delete CallMonitorChild records
			UPDATE CallMonitorChild SET LogicalDelete = 1 WHERE ActivityID = @ActivityID
			IF @@ERROR <> 0
			BEGIN
			    	ROLLBACK
				SELECT Results = 'Error: Cannot logically delete CallMonitorChild data'
			    	RETURN
			END
	
			-- // Logically delete CallMonitorParent record
			UPDATE CallMonitorParent SET LogicalDelete = 1 WHERE ActivityID = @ActivityID
			IF @@ERROR <> 0
			BEGIN
			    	ROLLBACK
				SELECT Results = 'Error: Cannot logically delete CallMonitorParent record'
			    	RETURN
			END

		END
		ELSE IF @Delete = 0
		BEGIN

			-- // Logically undelete CallMonitorChild records
			UPDATE CallMonitorChild SET LogicalDelete = 0 WHERE ActivityID = @ActivityID
			IF @@ERROR <> 0
			BEGIN
			    	ROLLBACK
				SELECT Results = 'Error: Cannot logically undelete CallMonitorChild data'
			    	RETURN
			END
	
			-- // Logically delete CallMonitorParent record
			UPDATE CallMonitorParent SET LogicalDelete = 0 WHERE ActivityID = @ActivityID
			IF @@ERROR <> 0
			BEGIN
			    	ROLLBACK
				SELECT Results = 'Error: Cannot logically undelete CallMonitorParent record'
			    	RETURN
			END


		END






	END
		
COMMIT

SELECT Results = 'Success'







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






/*
DECLARE @RelationCode varchar(20)
DECLARE @InsuredName varchar(52)
DECLARE @Tier varchar(50)
DECLARE @ProdDetailID varchar(20)
exec ProjectReports..usp_GetConditionedFields '20100203145131753', @ProdDetailID OUTPUT, @RelationCode OUTPUT, @InsuredName OUTPUT, @Tier OUTPUT
Select ProdDetailID = @ProdDetailID, RelationCode = @RelationCode, InsuredName = @InsuredName, Tier = @Tier
*/

CREATE       Procedure usp_GetConditionedFields
(
@AppID varchar(30),
@ProdDetailID varchar(25) OUTPUT,
@RelationCode varchar(20) OUTPUT,
@InsuredName varchar(52) OUTPUT, 
@Tier varchar(50) OUTPUT
)

AS

BEGIN

-- // Final AcesSpecial conversion for family products will determine default value upon discovery of ClientSubProductID

SELECT
@ProdDetailID = aps.ProductID,

@RelationCode = case
	when aps.RelationCode is null or aps.RelationCode = '' then 'E'
	Else ProjectReports.dbo.ufn_GetPreferredAPSRelationCode(aps.RelationCode, 0, 1)
	End,
@InsuredName = case
	when aps.InsuredMI is null then aps.InsuredLastName + ', ' + aps.InsuredFirstName
	else aps.InsuredLastName + ', ' + aps.InsuredFirstName + ' ' + aps.InsuredMI
	end,

@Tier = ProjectReports.dbo.ufn_GetTier(@AppID)

FROM IAMS..AppsAndPolsSummary aps WHERE aps.AppID = @AppID

END







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO








/*
  exec usp_GetDailyHoursWorked 
@EnrollerID =  null , @Location = 'HBG', @StartDate = '10/17/2011', @EndDate = '10/20/2011'
*/

CREATE    PROCEDURE [dbo].[usp_GetDailyHoursWorked]
(
 @EnrollerID varchar(50),
 @Location varchar(50),
 @StartDate datetime,
 @EndDate datetime
)
---------------------------------------------------------------
-- NAME: GetDailyHoursWorked
--
-- PURPOSE: Daily enroller work hours report
--
-- INPUTS: 
--
-- OUTPUTS: 
--
-- HISTORY:
--	10/20/11 - RMB - Created

----------------------------------------------------------------

AS
SET NOCOUNT ON


--SELECT @EnrollerID=NULL		--NULL = all
--SELECT @Location=NULL		--NULL = both, or LA/HBG
--SELECT @StartDate='1/1/2011'
--SELECT @EndDate='12/31/2011'


/* Old query:
Select 
	Enroller=rpt.EnrollerID,
	u.LocationID,
	TotalHour=SUM(TotalHours),
	EnrollmentHours=SUM(EnrollHours),
	AdminHours=SUM(AdminHours),
	CoachingHours=SUM(CoachHours),
	TrainingHours=SUM(TrainHours),
	StartDate=CONVERT(varchar(100),Min(Rpt.CallDate),101),
	EndDate=CONVERT(varchar(100),Max(Rpt.CallDate),101),
	EnrollerName = u.LastName  + ', ' + u.FirstName
From UserManagement..Users as U
INNER JOIN Rpt_CallHistory as rpt
ON
	rpt.EnrollerID=U.UserID
	AND (@Location=u.LocationID or @Location is NULL)
	AND (@EnrollerID = u.UserID or @EnrollerID is NULL)
	AND u.ROLE in ('ENROLLER','SUPERVISOR') 
    	AND u.CompanyID='BVI'
	AND rpt.CallDate between @StartDate and @EndDate
Group By rpt.EnrollerID, u.locationID --, u.LastName, u.FirstName
Order by u.LocationID, rpt.EnrollerID --, u.LastName, u.FirstName
--Group By EnrollerName, u.locationID
--Order by u.LocationID, EnrollerName
*/

Select 
	Enroller = edp.EnrollerID,
	EnrollerName = u.LastName  + ', ' + u.FirstName,
	u.LocationID,
	TotalHours=SUM(TotalHours),
	EnrollmentHours=SUM(EnrollHours),
	AdminHours=SUM(AdminHours),
	CoachingHours=SUM(CoachHours),
	TrainingHours=SUM(TrainHours),
	StartDate=CONVERT(varchar(100),Min(edp.ProjectDate),101),
	EndDate=CONVERT(varchar(100),Max(edp.ProjectDate),101)
From UserManagement..Users as U
INNER JOIN ProjectReports..EnrollerDateProject as edp
ON
	edp.EnrollerID=U.UserID
	AND (@Location=u.LocationID or @Location is NULL)
	AND (@EnrollerID = u.UserID or @EnrollerID is NULL)
	AND u.ROLE in ('ENROLLER','SUPERVISOR') 
    	AND u.CompanyID='BVI'
	AND edp.ProjectDate between @StartDate and @EndDate
--Group By edp.EnrollerID, u.locationID, u.LastName, u.FirstName
--Order by u.LocationID, edp.EnrollerID, u.LastName, u.FirstName
Group By  u.LocationID, u.LastName, u.FirstName, edp.EnrollerID
Order by u.LocationID, u.LastName, u.FirstName, edp.EnrollerID




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE   PROCEDURE usp_GetDatesUpdateRequired
AS

BEGIN

	DECLARE @Recordcount int
	DECLARE @DateLastUpdate datetime

	SELECT @Recordcount = Count (*) FROM Rpt_Control
	IF @Recordcount = 0 
	Begin
		SELECT DISTINCT NeedsUpdate = Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 
		FROM EmpTransmittal 
		ORDER  BY  Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 
	End
	Else
	Begin

SELECT @DateLastUpdate = MAX(DateLastUpdate) FROM Rpt_Control

		SELECT DISTINCT NeedsUpdate = Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 
		FROM EmpTransmittal 
		WHERE ChangeDate > @DateLastUpdate
		ORDER  BY  Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 
	End
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE   PROCEDURE  dbo.usp_GetExcelData (
	@ClientID as varchar(20),
	@EnrollerID as varchar (20),
	@CallStartTime as datetime
)

AS


/*
Declare @ClientID varchar(20)
Declare @EnrollerID varchar(20)
Declare @CallStartTime varchar(10)
Select @ClientID = 'Genesis'
Select @EnrollerID = 'lbrogan'
--Select @CallStartTime = Cast('10/20/2008' as datetime)
Select @CallStartTime = '10/20/2008'
*/



Declare @Sql varchar(500)

-- 1. Get a list of EmpIDs for this client for this date.

IF OBJECT_ID('ProjectReports..#Emps') IS NOT NULL  DROP TABLE #Emps

CREATE Table #Emps (EmpID int)

SELECT @Sql = 'INSERT INTO #Emps (EmpID) '
SELECT @Sql = @Sql + 'SELECT DISTINCT et.EmpID FROM EmpTransmittal et '
SELECT @Sql = @Sql + 'INNER JOIN ' + @ClientID  + '..Employee e on et.EmpID = e.EmpID '
SELECT @Sql = @Sql + 'WHERE et.EnrollerID = ''' + @EnrollerID + ''' AND dbo.ufn_IsDateEqual(et.callstarttime, ''' + @CallStartTime + ''')=1'

--Select dog =  @Sql
exec (@Sql)
--Select * FROM #Emps

SELECT tm.EmpID, 
Name = e.LastName + ', ' + e.FirstName,
NewHires =
	 (
	SELECT Count (*) FROM ProjectReports..EmpTransmittal 
	WHERE EnrollerID = @EnrollerID AND 
	EmpID = tm.EmpID AND
	EnrollWinCode = 'NH' AND 
	ActivityTypeCode <> 'INCOMPLETE' AND
	dbo.ufn_IsDateEqual(ProjectReports..EmpTransmittal.callstarttime, @CallStartTime) =1 
	 ),
OpenEnrollment = 
	(
	SELECT Count (*) FROM ProjectReports..EmpTransmittal 
	WHERE EnrollerID = @EnrollerID AND 
	EmpID = tm.EmpID AND
	EnrollWinCode = 'OE' AND 
	ActivityTypeCode <> 'INCOMPLETE' AND
	dbo.ufn_IsDateEqual(ProjectReports..EmpTransmittal.callstarttime, @CallStartTime) =1 
	),
LSCEnrollment = 
	(
	SELECT Count (*) FROM ProjectReports..EmpTransmittal 
	WHERE EnrollerID = @EnrollerID AND 
	EmpID = tm.EmpID AND
	EnrollWinCode = 'SC' AND 
	ActivityTypeCode <> 'INCOMPLETE' AND
	dbo.ufn_IsDateEqual(ProjectReports..EmpTransmittal.callstarttime, @CallStartTime) =1 
	),
CSRCalls = 
	(
	SELECT Count (*) FROM ProjectReports..EmpTransmittal 
	WHERE EnrollerID = @EnrollerID AND 
	EmpID = tm.EmpID AND
	EnrollWinCode = 'CSR' AND 
	ActivityTypeCode <> 'INCOMPLETE' AND
	dbo.ufn_IsDateEqual(ProjectReports..EmpTransmittal.callstarttime, @CallStartTime) =1 
	),
EnrollHours =
	(
	SELECT IsNull(SUM(EnrollHours), 0) FROM ProjectReports..EnrollerDateProject
	WHERE EnrollerID = @EnrollerID AND 
	dbo.ufn_IsDateEqual(ProjectDate, @CallStartTime) =1 
	),
AdminHours =
	(
	SELECT IsNull(SUM(AdminHours), 0) FROM ProjectReports..EnrollerDateProject
	WHERE EnrollerID = @EnrollerID AND 
	dbo.ufn_IsDateEqual(ProjectDate, @CallStartTime) =1 
	),
TotalHours =
	(
	SELECT IsNull(SUM(AdminHours), 0) + IsNull(SUM(AdminHours), 0) FROM ProjectReports..EnrollerDateProject
	WHERE EnrollerID = @EnrollerID AND 
	dbo.ufn_IsDateEqual(ProjectDate, @CallStartTime) =1 
	)

FROM #Emps tm
INNER JOIN Genesis..Employee e on tm.EmpID = e.EmpID
ORDER BY tm.EmpID

Drop Table #Emps



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


/*

To: plebi@allstate.com
cc: jsnavely@benefitvision.com
subj: GNC production report 10/27/2010 - 11/12/2010
Run every Monday

Date range: 10/27/2010 to last work day.




usp_JonSnavelyDiopitt  '12/14/2010'
*/


CREATE        PROCEDURE  usp_JonSnavelyDiopitt
(
	@EndDate datetime
)	

AS


SET NOCOUNT ON


SELECT
 [Call Date] = et.CallStartTime, 
[Emp Name] =et.LastName + ', ' + et.FirstName, 
Placeholder = '',

Enrolled = case
when ept.ProductID IS NULL then 0
else 1
end,

Product = ISNULL(ept.ProductID, ''),

Tier = case 
when (ept.ProductID = 'ALLSTATEUL') AND
(alt.AltProductDataID IS NULL) AND
(left(apd.FieldName, 22) = 'Future Purchase Option') then 'EZ' 
when (ept.ProductID = 'TMARKCOMBO')  AND
(alt.AltProductDataID IS NULL) AND
(apd.FieldName='EZValue' AND
apd.FieldData = 'true') then 'EZ1_5' 
when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV1', apd.FieldData) > 0) then 'EZ1_5' 
when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV2', apd.FieldData) > 0) then 'EZ1_10' 
when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV3', apd.FieldData) > 0) then 'EZ2_5' 
when (ept.ProductID = 'TRANSUL')   AND
(alt.AltProductDataID IS NULL) AND
(apd.FieldName = 'EZValue') then 'EZ' 
else ISNULL(alt.Tier, '') end ,

WeeklyPremium = case when ccp.DisplayPremiumInd = 1 then Cast(ept.WeeklyPremium as money) else null end,
AnnualPremium = case when ccp.DisplayPremiumInd = 1 then Cast(ept.AnnualPremium as money) else null end, 
PlanBenefitAmt = case when ccp.DisplayPlanBenefitAmtInd = 1 then Cast(ept.PlanBenefitAmt as money) else null end

FROM EmpTransmittal et
LEFT JOIN EmpProductTransmittal ept on et.activityID = ept.activityID
LEFT JOIN Config_ClientProduct ccp ON ept.ProductID = ccp.ClientProductID 
LEFT JOIN IAMS..AppsAndPolsSummary aps ON aps.AppID = ept.AppID AND aps.ActivityID = et.ActivityID 

LEFT JOIN IAMS..AppsAndPolsData apd ON aps.AppID = apd.AppID AND
((aps.ProductID = 'ALLSTATEUL' AND
apd.FieldName = 'Future Purchase Option Rider') OR (aps.ProductID IN ('TRANSUL', 'TMARKCOMBO') AND
apd.FieldName='EZValue' AND
apd.FieldData = 'true') OR (aps.ProductID = 'TMARKUL'  AND
CharIndex('EZV', apd.FieldData) > 0) ) 

LEFT JOIN Alt_ProductData alt ON alt.AltProductDataID = ept.AltProductDataID 

where et.clientid= 'Diopitt' and 
dbo.ufn_IsDateBetween(et.CallStartTime, '10/27/2010', @EndDate) = 1

-- added 12/14/2010
and isnull(et.LogicalDelete, 0) <> 1 and isnull(ept.LogicalDelete, 0) <> 1 and et.SupervisorApprovalDate IS NOT NULL


AND
(
                (aps.ActivityID = ept.ActivityID AND aps.AppID = ept.AppID AND aps.BVIAppStatus <> 'CANCELLED')
                OR
               (aps.ActivityID IS NULL AND aps.AppID IS NULL AND alt.AltProductDataID IS NOT NULL)
)

order by et.Lastname + et.Firstname

SET NOCOUNT ON








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/*

To: plebi@allstate.com
cc: jsnavely@benefitvision.com
subj: GNC production report 10/27/2010 - 11/12/2010
Run every Monday

Date range: 10/27/2010 to last work day.




usp_JonSnavely '11/14/2010'
*/


CREATE         PROCEDURE  usp_JonSnavelyGNC
(
	@EndDate datetime
)	

AS


SET NOCOUNT ON


SELECT
 [Call Date] = et.CallStartTime, 
[Emp Name] =et.LastName + ', ' + et.FirstName, 
Placeholder = '',

Enrolled = case
when ept.ProductID IS NULL then 0
else 1
end,

Product = ISNULL(ept.ProductID, ''),

Tier = case 
when (ept.ProductID = 'ALLSTATEUL') AND
(alt.AltProductDataID IS NULL) AND
(left(apd.FieldName, 22) = 'Future Purchase Option') then 'EZ' 
when (ept.ProductID = 'TMARKCOMBO')  AND
(alt.AltProductDataID IS NULL) AND
(apd.FieldName='EZValue' AND
apd.FieldData = 'true') then 'EZ1_5' 
when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV1', apd.FieldData) > 0) then 'EZ1_5' 
when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV2', apd.FieldData) > 0) then 'EZ1_10' 
when (ept.ProductID = 'TMARKUL')  AND
(alt.AltProductDataID IS NULL) AND
(CharIndex('EZV3', apd.FieldData) > 0) then 'EZ2_5' 
when (ept.ProductID = 'TRANSUL')   AND
(alt.AltProductDataID IS NULL) AND
(apd.FieldName = 'EZValue') then 'EZ' 
else ISNULL(alt.Tier, '') end ,

WeeklyPremium = case when ccp.DisplayPremiumInd = 1 then Cast(ept.WeeklyPremium as money) else null end,
AnnualPremium = case when ccp.DisplayPremiumInd = 1 then Cast(ept.AnnualPremium as money) else null end, 
PlanBenefitAmt = case when ccp.DisplayPlanBenefitAmtInd = 1 then Cast(ept.PlanBenefitAmt as money) else null end

FROM EmpTransmittal et
LEFT JOIN EmpProductTransmittal ept on et.activityID = ept.activityID
LEFT JOIN Config_ClientProduct ccp ON ept.ProductID = ccp.ClientProductID 
LEFT JOIN IAMS..AppsAndPolsSummary aps ON aps.AppID = ept.AppID AND aps.ActivityID = et.ActivityID 

LEFT JOIN IAMS..AppsAndPolsData apd ON aps.AppID = apd.AppID AND
((aps.ProductID = 'ALLSTATEUL' AND
apd.FieldName = 'Future Purchase Option Rider') OR (aps.ProductID IN ('TRANSUL', 'TMARKCOMBO') AND
apd.FieldName='EZValue' AND
apd.FieldData = 'true') OR (aps.ProductID = 'TMARKUL'  AND
CharIndex('EZV', apd.FieldData) > 0) ) 

LEFT JOIN Alt_ProductData alt ON alt.AltProductDataID = ept.AltProductDataID 

where et.clientid= 'GNC'  and 
dbo.ufn_IsDateBetween(et.CallStartTime, '10/27/2010', @EndDate) = 1

-- added 12/14/2010
and isnull(et.LogicalDelete, 0) <> 1 and isnull(ept.LogicalDelete, 0) <> 1 and et.SupervisorApprovalDate IS NOT NULL

order by et.Lastname + et.Firstname

SET NOCOUNT ON









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/*
exec usp_ProjectReportsMigrate_APSDataChanged
select  * from DataMigration_APSDataChanged order by adddate
*/

CREATE    PROCEDURE  usp_ProjectReportsMigrate_APSDataChanged
AS
SET NOCOUNT ON


IF EXISTS 
(SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'
AND TABLE_NAME = 'DataMigration_APSDataChanged')
DROP TABLE DataMigration_APSDataChanged

SELECT 
ept.AppID,
ept.ActivityID,
ept.AddDate,
ept.ClientID,
ProductID = ccp.BVIProductID,
EPT_PerPayPremium= ept.PerPayPremium, 
APS_PerPayPremium = aps.PerPayPremium,
EPT_PlanBenefitAmt = ept.PlanBenefitAmt,
APS_PlanBenefitAmt = aps.PlanBenefitAmt,
et.PayFrequencyCode

INTO DataMigration_APSDataChanged
FROM EmpProductTransmittal ept 
INNER JOIN EmpTransmittal et ON ept.ActivityID = et.ActivityID
LEFT JOIN Config_ClientProduct ccp on ept.productid = ccp.clientproductid
INNER JOIN IAMS..AppsAndPolsSummary aps on ept.ActivityID = aps.ActivityID and ept.AppID = aps.AppID

WHERE 
ept.LogicalDelete <> 1 AND 
(
(ept.PerPayPremium < aps.PerPayPremium * .9) OR (ept.PerPayPremium  > aps.PerPayPremium * 1.1) OR
(ept.PlanBenefitAmt < aps.PlanBenefitAmt * .9) OR (ept.PlanBenefitAmt  > aps.PlanBenefitAmt * 1.1)
)
  

SET NOCOUNT ON




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/*
exec usp_ProjectReportsMigrate_NoAPSRecord
select  * from DataMigration_NoAPSRecord
*/

CREATE     PROCEDURE  usp_ProjectReportsMigrate_NoAPSRecord
AS
SET NOCOUNT ON

IF EXISTS 
(SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'
AND TABLE_NAME = 'DataMigration_NoApsRecord')
DROP TABLE DataMigration_NoAPSRecord

SELECT 
ept.ClientID, 
ProductID = ccp.BVIProductID,
ept.AddDate,
ept.EmpID, 
alt.InsuredName,
alt.RelationCode,
ept.PerPayPremium, 
ept.PlanBenefitAmt,
et.PayFrequencyCode,
NonAcesPayFrequencyDesc = cpfna.PayFrequencyDesc,
ept.LicensedEnroller, 
ept.ActivityID, 
alt.Tier

INTO DataMigration_NoAPSRecord
FROM EmpProductTransmittal ept 
INNER JOIN EmpTransmittal et ON ept.ActivityID = et.ActivityID
LEFT JOIN Config_ClientProduct ccp on ept.productid = ccp.clientproductid
LEFT JOIN IAMS..AppsAndPolsSummary aps ON  ept.AppID = aps.AppID
LEFT JOIN Alt_ProductData alt on ept.AltProductDataID = alt.AltProductDataID
LEFT JOIN Codes_PayFrequency_NonAces cpfna ON ept.ClientID = cpfna.ClientID and et.PayFrequencyCode = cpfna.PayFrequencyCode
WHERE 
ept.LogicalDelete <> 1 AND
aps.ActivityID IS NULL

SET NOCOUNT ON



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/*
exec usp_ProjectReportsMigrate_NoAPSRecord
select  * from DataMigration_NoAPSRecord
*/

CREATE     PROCEDURE  [dbo].[usp_ProjectReportsMigrate_NoAPSRecord_Mike]
AS
SET NOCOUNT ON

IF EXISTS 
(SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'
AND TABLE_NAME = 'DataMigration_NoApsRecord')
DROP TABLE DataMigration_NoAPSRecord_Mike

SELECT 
ept.ClientID, 
ProductID = ccp.BVIProductID,
ept.AddDate,
ept.EmpID, 
ept.AppID,
--alt.InsuredName,
--alt.RelationCode,
ept.PerPayPremium, 
ept.PlanBenefitAmt,
et.PayFrequencyCode,
--NonAcesPayFrequencyDesc = cpfna.PayFrequencyDesc,
ept.LicensedEnroller, 
ept.ActivityID
--,alt.Tier

INTO DataMigration_NoAPSRecord_Mike
FROM EmpProductTransmittal ept 
INNER JOIN EmpTransmittal et 
ON 
	ept.ActivityID = et.ActivityID and ept.LogicalDelete <> 1 
	and ept.ProductID <> 'EYEMEDVIS'
-- We do not get comission on Eyemed anymore, no need to track it.
	and ept.clientID not in ('GNC','CHOICES','COKC','Superior','OPTIONS','ARMSTRONG')
-- We know the above will need to be put in, they can be a stand-alone query, we know they won't exist in the IAMS DB.

--LEFT JOIN Alt_ProductData alt on ept.AltProductDataID = alt.AltProductDataID
--LEFT JOIN Codes_PayFrequency_NonAces cpfna ON ept.ClientID = cpfna.ClientID and et.PayFrequencyCode = cpfna.PayFrequencyCode
--We can do the above as a second query later
INNER JOIN Config_ClientProduct ccp on ept.productid = ccp.clientproductid
LEFT JOIN IAMS..log_AppsAndPolsSummary aps ON  
	
-- ept.AppID = aps.AppID -- This concerns me.  This isn't how I'd check for IAMS.   Need to check on Product Code / Date instead.   Also need to use log_ not the main table incase the app ever got cancelled/etc...
		(CONVERT(varchar(20),ept.Adddate,101) = CONVERT(varchar(20),aps.Adddate,101)
		and (ept.ProductID = aps.ProductID or (ept.ProductID='AMGENCANCER' and aps.ProductID='AIGCCI'))
		and ept.EmpID = aps.EmpID
		and aps.BVIAppStatus in ('NEW','POLICY') 
		and ept.ActivityID = aps.ActivityID)
	or
		ept.AppID = aps.AppID
WHERE 
aps.ActivityID IS NULL

-- Now we can add all the ones we KNOW will be missing:
SELECT 
ept.ClientID, 
ProductID = ccp.BVIProductID,
ept.AddDate,
ept.EmpID, 
ept.AppID,
alt.InsuredName,
alt.RelationCode,
ept.PerPayPremium, 
ept.PlanBenefitAmt,
et.PayFrequencyCode,
NonAcesPayFrequencyDesc = cpfna.PayFrequencyDesc,
ept.LicensedEnroller, 
ept.ActivityID
,alt.Tier

INTO DataMigration_NoAPSRecord_Mike
FROM EmpProductTransmittal ept 
INNER JOIN EmpTransmittal et 
ON 
	ept.ActivityID = et.ActivityID and ept.LogicalDelete <> 1 
	and ept.clientID in ('GNC','CHOICES','COKC','Superior','OPTIONS','ARMSTRONG')
INNER JOIN Alt_ProductData alt on ept.AltProductDataID = alt.AltProductDataID
INNER JOIN Codes_PayFrequency_NonAces cpfna ON ept.ClientID = cpfna.ClientID and et.PayFrequencyCode = cpfna.PayFrequencyCode
INNER JOIN Config_ClientProduct ccp on ept.productid = ccp.clientproductid

SET NOCOUNT ON



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    procedure usp_QADC_Verify_EmpTransmittal_vs_ClientCallActivity 
as

declare @startdate as smalldatetime
set @startdate=dateadd(dd, -5, getdate())

/*case
		when datediff(dd, (select cast(field1 as datetime) from qaclientfeatures where categoryid='emptransmittal'), getdate())<1 then 
			dateadd(dd, case
				when DATEPART(WEEKDAY, GETDATE())='1' then -3
				when DATEPART(WEEKDAY, GETDATE())='2' then -4 
				else -2 end, getdate())
		else (select cast(field1 as datetime) from qaclientfeatures where categoryid='emptransmittal') end

*/  --to be fixed when updating logic is fixed.		

--select DATEADD(hh, -5, GETUTCDATE()) as CurrTime
select @startdate as 'date used', isnull (client.clientid, project.clientid) as ClientID, client.EmpID as ClientEmpID, client.adddate as 'Empactivitylog Date', 
	case 
	when  client.activityid is null then 'Orphan Record'
	when  project.activityid is null then 'Missing Record'
	when  client.activityid is not null and project.activityid is not null and isnull (project.empid, '') <> isnull (client.empid, '') then 'Empid'
--	when client.activityid is not null and project.activityid is not null and isnull (project.lastname, '') <> isnull (client.lastname, '') then 'Lastname'
--	when client.activityid is not null and project.activityid is not null and isnull (project.firstname, '') <> isnull (client.firstname, '') then 'Firstname'
--	when client.activityid is not null and project.activityid is not null and isnull (project.payfrequencycode, '') <> isnull (client.payfrequencycode, '') then 'Payfrequencycode'
--	when client.activityid is not null and project.activityid is not null and isnull (project.enrollwincode, '') <> isnull (client.enrollwincode, '') then 'Enrollwincode'
	when client.activityid is not null and project.activityid is not null and isnull (project.confirmationstatement, '') <> isnull (client.confirmfilename, '') then 'confirmationfilename'
	when client.activityid is not null and project.activityid is not null and isnull (project.notation, '') <> isnull (client.notation, '') then 'Notation'
--	when client.activityid is not null and project.activityid is not null and isnull (project.callstarttime, '') <> isnull (client.callstarttime, '') then 'Callstarttime'
--	when client.activityid is not null and project.activityid is not null and isnull (project.callendtime, '') <> isnull (client.callendtime, '') then 'Callendtime' 
	else ''
	end as 'Error Reason', 
	EC, client.adddate as EAAddDt, client.callstarttime as EAChgDt, project.adddate as CCMAddDt, 
	client.activityid as ClientActID, client.enrollerid as ClientEnrollerID, client.enrollwincode as ClientEnrollWinCode, client.activitytypecode as ClientActType, client.notation as ClientNote, 
	project.activityid as CCMActID, project.EmpID as CCMEmpID, project.enrollerid as CCMEnrollerID, client.enrollwincode as CCMEnrollWin, client.activitytypecode as CCMActType, project.notation as CCMNote, 
	client.confirmfilename as EAConf, project.confirmationstatement as CCMConfirm
, isnull (project.empid,''), isnull (client.empid,''), isnull (project.confirmationstatement,''), isnull (client.confirmfilename,''), isnull (project.notation,'') , isnull (client.notation,'')
from 
(select 'BureauVeritas' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC,  
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from bureauveritas..empactivitylog ea 
		join bureauveritas..employee ee on 
			ea.empid = ee.empid
 	where ea.activitytypecode = 'call' and ea.notation like '%quit:%' 
		and ea.callstarttime >= @startdate
union
select 'CTCA' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from ctca..empactivitylog ea 
		join ctca..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate
union
select 'Fulton' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from Fulton..empactivitylog ea 
		join Fulton..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
		and ea.callstarttime >= @startdate
union
select 'Genesis' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from Genesis..empactivitylog ea 
		join Genesis..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate
union
select 'HT' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from HT..empactivitylog ea 
		join HT..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate
union
select 'MartinRea' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from MartinRea..empactivitylog ea 
		join MartinRea..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate
union
select 'Morgans' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from Morgans..empactivitylog ea 
		join Morgans..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate
union
select 'PKOH' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from PKOH..empactivitylog ea 
		join PKOH..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate
union
select 'RDS' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from RDS..empactivitylog ea 
		join RDS..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate
union
select 'Stantec' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate  
	from Stantec..empactivitylog ea 
		join Stantec..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate
union
select 'Weathershield' as ClientID, ea.activityid, ea.empid, ea.enrollerid, ee.lastname, ee.firstname, ee.payfrequencycode, ee.ssn, ea.enrollwincode, ea.enrollercomputername as EC, 
		ea.activitytypecode, ea.confirmfilename, ea.notation, ea.callstarttime, ea.callendtime, ea.adddate, ea.changedate 
	from Weathershield..empactivitylog ea 
		join Weathershield..employee ee on 
			ea.empid = ee.empid
 	where activitytypecode = 'call' and notation like '%quit:%' 
	and ea.callstarttime >= @startdate) client
full outer join projectreports..emptransmittal project on
	client.clientid = project.clientid and project.clientid in ('bureauveritas', 'CTCA', 'Fulton', 'Genesis', 'HT', 'Martinrea', 'morgans', 'RDS', 'stantec', 'Weathershield') and 
	client.empid = project.empid and 
	client.activityid = project.activityid
	
where 	client.callstarttime>'06/03/2009 14:50' and project.callstarttime>'06/03/2009 14:50'
and ((client.activityid is null and project.activityid is not null and project.adddate >= @StartDate) or 
	(client.activityid is not null and project.activityid is null) or
	(client.activityid is not null and project.activityid is not null and project.adddate >= @StartDate and
	 (isnull (project.empid,'') <> isnull (client.empid,'')
	or isnull (project.empid,'') <> isnull (client.empid,'')
--	or project.lastname <> client.lastname
--	or project.firstname <> client.firstname
--	or project.payfrequencycode <> client.payfrequencycode
-- 	or project.enrollwincode <> client.enrollwincode
	or isnull (project.confirmationstatement,'') <> isnull (client.confirmfilename,'')
	or isnull (project.notation,'') <> isnull (client.notation,'')
	or isnull (project.confirmationstatement,'') <> isnull (client.confirmfilename,'')
	or isnull (project.notation,'') <> isnull (client.notation,'')
--	or project.callstarttime <> client.callstarttime
--	or project.callendtime <> client.callendtime
	)
	))
order by client.adddate 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


-- exec usp_RptBuildCalendar

CREATE   Procedure usp_RptBuildCalendar

AS
SET NOCOUNT ON

	-- // DATE TABLE
	Declare @DateUpdate Table(RecID int IDENTITY, DateUpdate smalldatetime)
	Declare @iYear int
	Declare @iMonth int
	Declare @iDay int
	Declare @DaysInMonth int
	Declare @ThisDate smalldatetime
	
	Declare @BeginDate smalldatetime
	Declare @EndDate smalldatetime
	
	Select @BeginDate = '12/20/2008'
	Select @EndDate = '3/26/2009'
	
	Declare @BeginYear int
	Declare @BeginMonth int
	Declare @BeginDay int
	
	Declare @EndYear int
	Declare @EndMonth int
	Declare @EndDay int
	
	Select @BeginYear = DatePart(year, @BeginDate)
	Select @BeginMonth = DatePart(month, @BeginDate)
	Select @BeginDay = DatePart(day, @BeginDate)

	Select @EndYear = DatePart(year, @EndDate)
	Select @EndMonth = DatePart(month, @EndDate)
	Select @EndDay = DatePart(day, @EndDate)

	Declare @Finished bit
	Select @Finished = 0

	Select @iYear = @BeginYear
	While @iYear <= 10000
	Begin
			If @iYear = @BeginYear
		Begin
			Select @iMonth = @BeginMonth
		End	
		Else
		Begin
			Select @iMonth = 1
		End
	
	
		While @iMonth < 10000
		Begin
			Select @DaysInMonth = dbo.ufn_GetDaysInMonth(@iYear, @iMonth)
			If @iYear = @BeginYear and @iMonth = @BeginMonth
			Begin
				Select @iDay = @BeginDay
			End	
			Else
			Begin
				Select @iDay = 1
			End
	
			While @iDay <= 10000
			Begin		
				Select @ThisDate = dbo.ufn_PadLeft(Cast(@iMonth as varchar(2)), 2, '0')   + '-' + dbo.ufn_PadLeft(Cast(@iDay as varchar(2)), 2, '0')+ '-'  + Cast(@iYear as varchar(4))
				INSERT INTO @DateUpdate  SELECT @ThisDate
				Select @iDay = @iDay + 1
				If @iYear = @EndYear and @iMonth = @EndMonth and @iDay > @EndDay
				Begin
					Select @Finished = 1
					Break;
				End
				Else
					If @iDay > @DaysInMonth
					Begin
						Break;
					End
			End
	
			If @Finished = 1
				Break;
	
			Select @iMonth = @iMonth + 1
	
			If @iYear = @EndYear and @iMonth > @EndMonth
			Begin
				Select @Finished = 1
				Break;
			End
			Else
			Begin
				If @iMonth > 12
					Break;
			End
		End

	
		If @Finished = 1
			Break;
	
		Select @iYear = @iYear + 1
		If @iYear > @EndYear
			Break;
	
	End

	--Select * FROM @DateUpdate
	Select ThisDate = Convert(varchar, DateUpdate, 101) FROM @DateUpdate

SET NOCOUNT ON






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO












/*
delete rpt_callhistory
select * from rpt_callhistory order by clientid
Select count (*) from rpt_callhistory
--select count (*) from rpt_producthistory
--Select * from rpt_producthistory

Declare @RetVal varchar(500)
Select @RetVal = '0|Success'
exec usp_RptBuildClientSegment1 '8/20/2009', '8/20/2009', 'SFWMD', @RetVal
Select RetVal = @RetVal
select count (*) from rpt_callhistory
select top 10 * from rpt_callhistory

*/

--    usp_BuildClientSegment 'Morgans', '(''HBG'')'
--    usp_BuildClientSegment 'Morgans', '(''OKC'')'
--    usp_BuildClientSegment 'Morgans', '(''HBG'', ''OKC'')'


CREATE   PROCEDURE usp_RptBuildClientSegment1(@CallDate smalldatetime, @AddDate datetime, @ClientID varchar(20), @RetVal varchar(500) OUTPUT)
AS

SET NOCOUNT ON

BEGIN

	-- // Declarations
	Declare @Err int
	Declare @Sql1 varchar(8000)
	Declare @Sql2 varchar(5000)
	Declare @FieldName varchar(20)
	Declare @FieldList varchar(200)
	Declare @EnhancedFieldList varchar(200)
	Declare @FieldTable TABLE(RecID int, FieldName varchar(20))
	Declare @RecID int
	Declare @Recordcount int

	---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	-- // FieldList and EnhancedFieldList
	SELECT @FieldList = Columns FROM Excel_SegmentConfigure WHERE SegmentId = @ClientID
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1050|usp_RptBuildClientSegment1: Error selecting Columns from Excel_SegmentConfigure'
		Return
	End
	SELECT @EnhancedFieldList = '(EnrollerID, ClientID, CallDate, AddDate, ' + Replace(@FieldList, '|', ',') + ') '
	

	-- // FieldTable 
	INSERT INTO @FieldTable Select * FROM dbo.ufn_GetTableFromList(@ClientID)
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1052|usp_RptBuildClientSegment1: Error inserting records into @FieldTable'
		Return
	End
	---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	-- // Delete the records that are about to be replaced
	DELETE Rpt_CallHistory WHERE ClientID = @ClientID AND dbo.ufn_IsDateEqual(CallDate, @CallDate) = 1
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1053|usp_RptBuildClientSegment1: Error deleting records from Rpt_CallHistory'
		Return
	End

	-- // Start the Sql
	SELECT @Sql1 = 'SELECT '
	SELECT @Sql1 = @Sql1 +  'EnrollerID = u.UserID,       '
	SELECT @Sql1 = @Sql1 + 'ClientID = ''' + @ClientID + ''',         '
	SELECT @Sql1 = @Sql1 + 'CallDate = ''' + Convert(varchar, @CallDate, 101) + ''',          '
	SELECT @Sql1 = @Sql1 + 'AddDate = ''' + Convert(varchar, @AddDate,  120) + ''',     '


	SELECT @RecID = 1
	SELECT @Recordcount = Count (*) FROM @FieldTable
	WHILE @RecID <= @Recordcount
	Begin
		SELECT @Sql2 = ''
		SELECT @FieldName = FieldName FROM @FieldTable WHERE RecID = @RecID

		exec usp_RptBuildClientSegment2 @CallDate, @ClientID, @FieldName, @Sql2 OUTPUT, @RetVal OUTPUT

		If @RetVal <> '0|Success'
		Begin
			Return
		End

		SELECT @Sql1 = @Sql1 + @Sql2
		SELECT @RecID = @RecID + 1
	End


	-- // Remove the trailing comma
	SELECT @Sql1 = Substring(@Sql1, 1, Len(@Sql1) - 1)

	-- // Complete the Sql
/*
	SELECT @Sql1 = @Sql1 + 'FROM UserManagement..Users u '
	SELECT @Sql1 = @Sql1 + 'WHERE '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpTransmittal etMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE etMain.LogicalDelete = 0 AND etMain.EnrollerID = u.UserID AND '
	--SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CurDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EnrollerDate edMain '
	SELECT @Sql1 = @Sql1 + 'WHERE edMain.EnrollerID = u.UserID AND dbo.ufn_IsDateEqual(edMain.ProjectDate, ''' + Cast(@CallDate as varchar(20)) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND edMain.SupervisorApproval = 1) >0 '
*/

	SELECT @Sql1 = @Sql1 + 'FROM UserManagement..Users u '
	SELECT @Sql1 = @Sql1 + 'WHERE '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpTransmittal etMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE etMain.LogicalDelete = 0 AND etMain.EnrollerID = u.UserID AND '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpProductTransmittal eptMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN ProjectReports..EmpTransmittal etMain on eptMain.ActivityID = etMain.ActivityID '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE eptMain.LogicalDelete = 0 AND eptMain.LicensedEnroller = u.UserID AND '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EnrollerDate edMain '
	SELECT @Sql1 = @Sql1 + 'WHERE edMain.EnrollerID = u.UserID AND dbo.ufn_IsDateEqual(edMain.ProjectDate, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND edMain.SupervisorApproval = 1) >0   '


	-- // Insert the records into Rpt_CallHistory
	SELECT @Sql1 = 'INSERT INTO Rpt_CallHistory ' + @EnhancedFieldList + @Sql1
	EXEC(@Sql1)
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1054|usp_RptBuildClientSegment1: Error inserting records into Rpt_CallHistory'
		Return
	End

END

SET NOCOUNT OFF

























GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO















/*
delete rpt_callhistory
select * from rpt_callhistory order by clientid
Select count (*) from rpt_callhistory
--select count (*) from rpt_producthistory
--Select * from rpt_producthistory

Declare @RetVal varchar(500)
Select @RetVal = '0|Success'
exec usp_RptBuildClientSegment101 '10/26/2010', '10/31/2010', 'C3', @RetVal
Select RetVal = @RetVal
select count (*) from rpt_callhistory
select top 10 * from rpt_callhistory

*/

--    usp_BuildClientSegment 'Morgans', '(''HBG'')'
--    usp_BuildClientSegment 'Morgans', '(''OKC'')'
--    usp_BuildClientSegment 'Morgans', '(''HBG'', ''OKC'')'


CREATE      PROCEDURE usp_RptBuildClientSegment101(@CallDate smalldatetime, @AddDate datetime, @ClientID varchar(20), @RetVal varchar(8000) OUTPUT)
AS

SET NOCOUNT ON

BEGIN

	-- // Declarations
	Declare @Err int
	Declare @Sql1 varchar(8000)
	Declare @Sql2 varchar(5000)
	Declare @FieldName varchar(20)
	Declare @FieldList varchar(200)
	Declare @EnhancedFieldList varchar(200)
	Declare @FieldTable TABLE(RecID int, FieldName varchar(20))
	Declare @RecID int
	Declare @Recordcount int

	---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	-- // FieldList and EnhancedFieldList
	SELECT @FieldList = Columns FROM Excel_SegmentConfigure WHERE SegmentId = @ClientID
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1050|usp_RptBuildClientSegment1: Error selecting Columns from Excel_SegmentConfigure'
		Return
	End
	SELECT @EnhancedFieldList = '(EnrollerID, ClientID, CallDate, AddDate, ' + Replace(@FieldList, '|', ',') + ') '
	

	-- // FieldTable 
	INSERT INTO @FieldTable Select * FROM dbo.ufn_GetTableFromList(@ClientID)
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1052|usp_RptBuildClientSegment1: Error inserting records into @FieldTable'
		Return
	End

	---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	-- // Delete the records that are about to be replaced
	DELETE Rpt_CallHistory WHERE ClientID = @ClientID AND dbo.ufn_IsDateEqual(CallDate, @CallDate) = 1
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1053|usp_RptBuildClientSegment1: Error deleting records from Rpt_CallHistory'
		Return
	End

	-- // Start the Sql
	SELECT @Sql1 = 'SELECT '
	SELECT @Sql1 = @Sql1 +  'EnrollerID = u.UserID,       '
	SELECT @Sql1 = @Sql1 + 'ClientID = ''' + @ClientID + ''',         '
	SELECT @Sql1 = @Sql1 + 'CallDate = ''' + Convert(varchar, @CallDate, 101) + ''',          '
	SELECT @Sql1 = @Sql1 + 'AddDate = ''' + Convert(varchar, @AddDate,  120) + ''',     '


	SELECT @RecID = 1
	SELECT @Recordcount = Count (*) FROM @FieldTable
	WHILE @RecID <= @Recordcount
	Begin
		SELECT @Sql2 = ''
		SELECT @FieldName = FieldName FROM @FieldTable WHERE RecID = @RecID

		exec usp_RptBuildClientSegment2 @CallDate, @ClientID, @FieldName, @Sql2 OUTPUT, @RetVal OUTPUT

		If @RetVal <> '0|Success'
		Begin
			Return
		End

		SELECT @Sql1 = @Sql1 + @Sql2
		SELECT @RecID = @RecID + 1
	End


	-- // Remove the trailing comma
	SELECT @Sql1 = Substring(@Sql1, 1, Len(@Sql1) - 1)

	-- // Complete the Sql
	SELECT @Sql1 = @Sql1 + 'FROM UserManagement..Users u '
	SELECT @Sql1 = @Sql1 + 'WHERE '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpTransmittal etMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE etMain.LogicalDelete = 0 AND etMain.EnrollerID = u.UserID AND '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpProductTransmittal eptMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN ProjectReports..EmpTransmittal etMain on eptMain.ActivityID = etMain.ActivityID '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE eptMain.LogicalDelete = 0 AND eptMain.LicensedEnroller = u.UserID AND '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EnrollerDate edMain '
	SELECT @Sql1 = @Sql1 + 'WHERE edMain.EnrollerID = u.UserID AND dbo.ufn_IsDateEqual(edMain.ProjectDate, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND edMain.SupervisorApproval = 1) >0   '


	-- // Insert the records into Rpt_CallHistory
	DELETE Rpt_CallHistoryWorking
	SELECT @Sql1 = 'INSERT INTO Rpt_CallHistoryWorking ' + @EnhancedFieldList + @Sql1
	EXEC(@Sql1)
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1054|usp_RptBuildClientSegment1: Error inserting records into Rpt_CallHistoryWorking'
		Return
	End

	INSERT INTO Rpt_CallHistory 
	SELECT * FROM Rpt_CallHistoryWorking WHERE
	IsNull(NewHires, 0) + IsNull(OpenEnrollment, 0) + IsNull(LSCEnrollment, 0) + IsNull(CSRCalls, 0) + IsNull(Interviewed, 0) + IsNull(Enrolled, 0) + IsNull(ServiceMode, 0) + IsNull(VolEnrollment, 0) + IsNull(BeneficiaryChange, 0) + IsNull(TotalHours, 0) > 0
	DELETE Rpt_CallHistoryWorking
	IF @Err <> 0
	Begin
		Select @RetVal = '1055|usp_RptBuildClientSegment1: Error inserting records into Rpt_CallHistory'
		Return
	End

END

SET NOCOUNT OFF




























GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

















/*

Declare @Sql2 varchar(5000)
Declare @RetVal varchar(500)
Declare @CallDate datetime
set @CallDate = '10/26/2010 '
SELECT @Sql2 = ''
 exec usp_RptBuildClientSegment2  @CallDate, 'C3', 'Enrolled', @Sql2 OUTPUT, @RetVal OUTPUT
SELECT @Sql2, @RetVal

*/
--    usp_BuildClientSegment 'Morgans', '(''HBG'')'
--    usp_BuildClientSegment 'Morgans', '(''OKC'')'
--    usp_BuildClientSegment 'Morgans', '(''HBG'', ''OKC'')'

CREATE          PROCEDURE usp_RptBuildClientSegment2(@CallDate smalldatetime, @ClientID varchar(20), @FldName varchar(20), @Sql2 varchar(5000) OUTPUT, @RetVal varchar(500) OUTPUT)
AS

SET NOCOUNT ON

BEGIN

	Declare @Err int
	Declare @EnrollWinCode varchar(12)
	Declare @Hours varchar(200)

	
	IF @FldName = 'Name'
		Set @Sql2 = @Sql2 + 'Name = u.LastName + '', '' + u.FirstName,  '
 
	ELSE IF @FldName IN ('NewHires', 'OpenEnrollment', 'VolEnrollment', 'LSCEnrollment', 'ServiceMode', 'BeneficiaryChange')
	Begin
		SELECT @EnrollWinCode = case
			when @FldName = 'NewHires' then 'NH'
			when @FldName = 'OpenEnrollment' then 'OE'
			when @FldName = 'VolEnrollment' then 'VE'
			when @FldName = 'LSCEnrollment' then 'SC'
			when @FldName = 'ServiceMode' then 'SM'
			when @FldName = 'BeneficiaryChange' then 'BC'
			end
	
		SELECT @Sql2 = @Sql2 +  @FldName + ' =  (SELECT Count (*) FROM ProjectReports..EmpTransmittal et2 '
		SELECT @Sql2 = @Sql2 + 'INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID '
		SELECT @Sql2 = @Sql2 +  dbo.ufn_GetClientJoin(@ClientID, 'et2')
		SELECT @Sql2 = @Sql2 + 'WHERE et2.LogicalDelete = 0 AND et2.EnrollWinCode = ''' +  @EnrollWinCode + ''' AND et2.ActivityTypeCode NOT IN (''INCOMPLETE'', ''CSR'') AND '
		SELECT @Sql2 = @Sql2 + dbo.ufn_GetClientWhereSelect(@ClientID, 'et2', 0)
		SELECT @Sql2 = @Sql2 + ' et2.EnrollerID = u.UserID AND '
		SELECT @Sql2 = @Sql2 +  'dbo.ufn_IsDateEqual(et2.CallStartTime, ''' + Convert(varchar, @CallDate, 101) + ''') = 1 AND '
		SELECT @Sql2 = @Sql2  + 'et2.SupervisorApprovalDate IS NOT NULL AND '
		SELECT @Sql2 = @Sql2  + 'dbo.ufn_IsTestID(bviclient.ClientID, et2.EmpID) = 0), '
	End

	-- // CSR activity type code overrides the enrollment window code
	ELSE IF @FldName = 'CSRCalls'
	Begin
		SELECT @Sql2 = @Sql2 + 'CSRCalls =  (SELECT Count (*) FROM ProjectReports..EmpTransmittal et2 '
		SELECT @Sql2 = @Sql2 + 'INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID '
		SELECT @Sql2 = @Sql2 +  dbo.ufn_GetClientJoin(@ClientID, 'et2')
		SELECT @Sql2 = @Sql2 + 'WHERE et2.LogicalDelete = 0 AND et2.EnrollWinCode IN (''NH'', ''OE'', ''VE'', ''SC'', ''CSR'', ''SM'', ''BC'')  AND '
		SELECT @Sql2 = @Sql2 + 'et2.ActivityTypeCode =  ''CSR'' AND '
		SELECT @Sql2 = @Sql2 + dbo.ufn_GetClientWhereSelect(@ClientID, 'et2', 0)
		SELECT @Sql2 = @Sql2 + ' et2.EnrollerID = u.UserID AND '
		SELECT @Sql2 = @Sql2 +  'dbo.ufn_IsDateEqual(et2.CallStartTime, ''' + Convert(varchar, @CallDate, 101) + ''') = 1 AND '
		SELECT @Sql2 = @Sql2  + 'et2.SupervisorApprovalDate IS NOT NULL AND '
		SELECT @Sql2 = @Sql2  + 'dbo.ufn_IsTestID(bviclient.ClientID, et2.EmpID) = 0),             '
	End


	ELSE IF @FldName = 'Interviewed'
	Begin



--		If @ClientID = 'C3'
--		Begin
--			SELECT @Sql2 = @Sql2 + 'Interviewed = 0, '
--		End
--		Else
--		Begin
			SELECT @Sql2 = @Sql2 +  @FldName + ' =  (SELECT Count (*) FROM ProjectReports..EmpTransmittal et2 '
			SELECT @Sql2 = @Sql2 + 'INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID '
			SELECT @Sql2 = @Sql2 +  dbo.ufn_GetClientJoin(@ClientID, 'et2')
			SELECT @Sql2 = @Sql2 + 'WHERE '
			IF @ClientID IN ('ht', 'weathershield')
			Begin
				SELECT @Sql2 = @Sql2 + 'EnrollWinCode <> ''SC'' AND EnrollWinCode <> ''SM'' AND '
			End
			SELECT @Sql2 = @Sql2 + 'EnrollWinCode <> ''BC'' AND '
			SELECT @Sql2 = @Sql2 + 'et2.LogicalDelete = 0 AND et2.ActivityTypeCode = ''CALL'' AND '
			SELECT @Sql2 = @Sql2 + dbo.ufn_GetClientWhereSelect(@ClientID, 'et2', 0)
			SELECT @Sql2 = @Sql2 + ' et2.EnrollerID = u.UserID AND '
			SELECT @Sql2 = @Sql2 +  'dbo.ufn_IsDateEqual(et2.CallStartTime, ''' + Convert(varchar, @CallDate, 101) + ''') = 1 AND '
			SELECT @Sql2 = @Sql2  + 'et2.SupervisorApprovalDate IS NOT NULL AND '
			SELECT @Sql2 = @Sql2  + 'dbo.ufn_IsTestID(bviclient.ClientID, et2.EmpID) = 0),             '
--		End



	End

	ELSE IF @FldName = 'Enrolled'
	Begin


--		If @ClientID = 'C3'
--		Begin
--			SELECT @Sql2 = @Sql2 + 'Enrolled = 0, '
--		End
--		Else
--		Begin

			SELECT @Sql2 = @Sql2 +  @FldName + ' = '
			SELECT @Sql2 = @Sql2 +  '('
			SELECT @Sql2 = @Sql2 +  'SELECT Count (*)  From EmpTransmittal etPrimary '	
			SELECT @Sql2 = @Sql2 +  'INNER JOIN '
			SELECT @Sql2 = @Sql2 +  '('
	
	                    	-- Non-IAMS: Extended
			SELECT @Sql2 = @Sql2 +  'SELECT DISTINCT ept2.ActivityID, ept2.ProductID '
			SELECT @Sql2 = @Sql2 +  'FROM EmpProductTransmittal ept2 '
			SELECT @Sql2 = @Sql2 +  'INNER JOIN BVI..Client bviclient ON ept2.ClientID = bviclient.ClientID '
			SELECT @Sql2 = @Sql2 +  'INNER JOIN ProjectReports..EmpTransmittal et2 ON ept2.ActivityID = et2.ActivityID '
			SELECT @Sql2 = @Sql2 +   dbo.ufn_GetClientJoin(@ClientID, 'et2')
			SELECT @Sql2 = @Sql2 +  'LEFT JOIN ClientProduct_Extended cpe on ept2.ClientId = cpe.ClientId AND ept2.ProductId = cpe.ClientProductID '
			SELECT @Sql2 = @Sql2 +  'WHERE ept2.LogicalDelete = 0 AND '
			SELECT @Sql2 = @Sql2 +   dbo.ufn_GetClientWhereSelect(@ClientID, 'et2', 0)
			SELECT @Sql2 = @Sql2 +  'cpe.ExtendedInd = 1 AND dbo.ufn_IsDateBetween(''' + Convert(varchar, @CallDate, 101) + ''', cpe.StartDate, cpe.EndDate) = 1 AND '
			SELECT @Sql2 = @Sql2 +  'ept2.LicensedEnroller = u.UserID AND '
			SELECT @Sql2 = @Sql2 +  'dbo.ufn_IsDateEqual(et2.CallStartTime, ''' + Convert(varchar, @CallDate, 101) + ''') = 1  AND '
			SELECT @Sql2 = @Sql2 +  'et2.SupervisorApprovalDate IS NOT NULL '
			
			SELECT @Sql2 = @Sql2 +  'UNION '
		
			-- IAMS: Standard and AcesSpecial
			SELECT @Sql2 = @Sql2 +  'SELECT DISTINCT ept2.ActivityID, ept2.ProductID '
			SELECT @Sql2 = @Sql2 +  'FROM ProjectReports..EmpProductTransmittal ept2 '
			SELECT @Sql2 = @Sql2 +  'INNER JOIN BVI..Client bviclient ON ept2.ClientID = bviclient.ClientID '
			SELECT @Sql2 = @Sql2 +  'INNER JOIN ProjectReports..EmpTransmittal et2 ON ept2.ActivityID = et2.ActivityID '
			SELECT @Sql2 = @Sql2 +  dbo.ufn_GetClientJoin(@ClientID, 'et2')

			-- 1/25/2011
			--SELECT @Sql2 = @Sql2 +  'INNER JOIN IAMS..AppsAndPolsSummary aps on ept2.ActivityID = aps.ActivityID and ept2.AppID = aps.AppID '
			SELECT @Sql2 = @Sql2 +  'INNER JOIN IAMS..AppsAndPolsSummary aps on ept2.AppID = aps.AppID '
			SELECT @Sql2 = @Sql2 +  'WHERE ept2.LogicalDelete = 0 AND '
			SELECT @Sql2 = @Sql2 +  dbo.ufn_GetClientWhereSelect(@ClientID, 'et2', 0)
			SELECT @Sql2 = @Sql2 +  'aps.BVIAppStatus <> ''CANCELLED'' AND '
			SELECT @Sql2 = @Sql2 +  'ept2.LicensedEnroller = u.UserID AND '
			SELECT @Sql2 = @Sql2 +  'dbo.ufn_IsDateEqual(et2.CallStartTime, ''' + Convert(varchar, @CallDate, 101) + ''') = 1 AND '
			SELECT @Sql2 = @Sql2 +  'et2.SupervisorApprovalDate IS NOT NULL '
			SELECT @Sql2 = @Sql2 +  ') '
			SELECT @Sql2 = @Sql2 +  'eptSecondary on etPrimary.ActivityID = eptSecondary.ActivityID '
			SELECT @Sql2 = @Sql2 +  '),              '

--		End

	End

	--ELSE IF @FldName = 'Enrolled'
	--Begin
	--	SELECT @Sql2 = @Sql2 +  @FldName + ' =  (SELECT Count (*) FROM ProjectReports..EmpTransmittal et2 '
	--	SELECT @Sql2 = @Sql2 + 'INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID '
	--	SELECT @Sql2 = @Sql2 +  dbo.ufn_GetClientJoin(@ClientID, 'et2')
	--	SELECT @Sql2 = @Sql2 + 'WHERE et2.LogicalDelete = 0 AND '
	--	SELECT @Sql2 = @Sql2 +  ' et2.EnrollerID = u.UserID AND '
	--	SELECT @Sql2 = @Sql2 +  dbo.ufn_GetClientWhere(@ClientID, 'et2')
	--	SELECT @Sql2 = @Sql2 +  'dbo.ufn_IsDateEqual(et2.CallStartTime, ''' + Convert(varchar, @CallDate, 101) + ''') = 1 AND '
	--	SELECT @Sql2 = @Sql2 + '(SELECT Count (*) FROM EmpProductTransmittal WHERE ActivityID = et2.ActivityID) > 0 AND '
	--	SELECT @Sql2 = @Sql2 + 'et2.SupervisorApprovalDate IS NOT NULL AND dbo.ufn_IsTestID(bviclient.ClientID, et2.EmpID) = 0),             '
	--End

	ELSE IF @FldName IN ('AdminHours', 'EnrollHours', 'TrainHours', 'CoachHours', 'TotalHours')
	Begin
	SELECT @Hours = case
		when @FldName = 'AdminHours' then ' IsNull(SUM(AdminHours), 0) '
		when @FldName = 'EnrollHours' then ' IsNull(SUM(EnrollHours), 0) '
		when @FldName = 'TrainHours' then ' IsNull(SUM(TrainHours), 0) '
		when @FldName = 'CoachHours' then ' IsNull(SUM(CoachHours), 0) '
		when @FldName = 'TotalHours' then ' IsNull(SUM(AdminHours), 0) + IsNull(SUM(EnrollHours), 0) + IsNull(SUM(TrainHours), 0) + IsNull(SUM(CoachHours), 0) '
		end
	
		SELECT @Sql2 = @Sql2 +  @FldName + ' = (SELECT ' + @Hours + ' FROM ProjectReports..EnrollerDateProject edp '
		SELECT @Sql2 = @Sql2 + 'INNER JOIN ProjectReports..EnrollerDate ed on edp.EnrollerID = ed.EnrollerID AND edp.ProjectDate = ed.ProjectDate '
		--SELECT @Sql2 = @Sql2 + 'WHERE edp.ClientID = ''' + @ClientID + ''' AND edp.EnrollerID = u.UserID AND '

		SELECT @Sql2 = @Sql2 + 'WHERE edp.EnrollerID = u.UserID AND '
		SELECT @Sql2 = @Sql2 +  dbo.ufn_GetClientWhereSelect(@ClientID, 'edp', 1) 
		SELECT @Sql2 = @Sql2 + 'dbo.ufn_IsDateEqual(ed.ProjectDate, ''' + Convert(varchar, @CallDate, 101) + ''') =1 AND ed.SupervisorApproval = 1),          '
	End

	SELECT @RetVal = '0|Success'
END

SET NOCOUNT OFF





















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO













/*
Select count (*) from rpt_callhistory
select count (*) from rpt_producthistory

delete rpt_callhistory
Delete rpt_producthistory

Declare @Sql2 varchar(5000)
Declare @RetVal varchar(500)
Declare @CallDate smalldatetime
Declare @AddDate datetime
Declare @ClientID varchar(20)
Declare @ProductID varchar(20)
Declare @FieldName varchar(20)
Set @Sql2 = 'EE = (SELECT Count (*) FROM  Alt_ProductData alt INNER JOIN EmpProductTransmittal ept on alt.AltProductDataID = ept.AltProductDataID INNER JOIN EmpTransmittal et2 on ept.ActivityID = et2.ActivityID INNER JOIN BVI..Client bviclient ON et2.ClientID = bviclient.ClientID WHERE ept.LogicalDelete = 0 AND et2.SupervisorApprovalDate IS NOT NULL AND et2.ClientID = ''HT'' AND  ept.LicensedEnroller = u.UserID AND ept.ProductID = ''AMGENCANCER'' AND alt.RelationCode = ''E'' AND dbo.ufn_IsDateEqual(et2.CallStartTime, ''09/02/2009'') = 1 AND dbo.ufn_IsTestID(bviclient.ClientID, et2.EmpID) = 0) '
Set @CallDate = '9/2/2009'
Set @ClientID = 'BureauVeritas'
Set @ProductID = 'AIGACC'
Set @FieldName = 'EE'
Set @AddDate = '8/20/2009'
Set @RetVal = '0|Success'
exec usp_RptBuildProductCompactor @CallDate, @AddDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT, @RetVal  OUTPUT
--Select @sql2
Select RetVal = @RetVal

*/


CREATE      PROCEDURE usp_RptBuildProductCompactor(@CallDate smalldatetime, @AddDate datetime,  @ClientID varchar(20), @ProductID varchar(20), @FieldName varchar(20), @Sql2 varchar(5000), @RetVal varchar(500) OUTPUT)
AS

SET NOCOUNT ON

	Declare @Err int
	Declare @EnrollerID varchar(20)
	Declare @FieldData decimal(18,2)
	Declare @Sql1 varchar(8000)
	Declare @Sql3 varchar(3000)


	Set @Sql2 = 'FieldData' + Right(@Sql2, len(@Sql2) - len(@FieldName))

	-- // Complete the main from clause of the sql
/*
	SELECT @Sql1 = 'FROM UserManagement..Users u '
	SELECT @Sql1 = @Sql1 + 'WHERE '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpTransmittal etMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE etMain.LogicalDelete = 0 AND etMain.EnrollerID = u.UserID AND '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EnrollerDate edMain '
	SELECT @Sql1 = @Sql1 + 'WHERE edMain.EnrollerID = u.UserID AND dbo.ufn_IsDateEqual(edMain.ProjectDate, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND edMain.SupervisorApproval = 1) >0            '
*/
--/*
	SELECT @Sql1 = 'FROM UserManagement..Users u '
	SELECT @Sql1 = @Sql1 + 'WHERE '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpTransmittal etMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE etMain.LogicalDelete = 0 AND etMain.EnrollerID = u.UserID AND '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpProductTransmittal eptMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN ProjectReports..EmpTransmittal etMain on eptMain.ActivityID = etMain.ActivityID '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on eptMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE eptMain.LogicalDelete = 0 AND eptMain.LicensedEnroller = u.UserID AND '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EnrollerDate edMain '
	SELECT @Sql1 = @Sql1 + 'WHERE edMain.EnrollerID = u.UserID AND dbo.ufn_IsDateEqual(edMain.ProjectDate, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND edMain.SupervisorApproval = 1) >0   '
--*/

	-- // Concatenate the selection and the from clause
	SELECT @Sql1 = @Sql2 + @Sql1 + '                        '

	-- // Reset Rpt_ProductHistoryWorking
	Delete Rpt_ProductHistoryWorking

	-- // Prepare the insert for Rpt_ProductHistoryWorking
	Set @Sql3 = 'INSERT INTO Rpt_ProductHistoryWorking (EnrollerID, CallDate, ClientID, ProductID, FieldName, FieldData)   '
	Set @Sql3 = @Sql3 + 'SELECT EnrollerID = u.UserID, CallDate = ''' + Convert(varchar, @CallDate, 101) + ''', '
	Set @Sql3 = @Sql3 + 'ClientID = ''' + @ClientID + ''', ProductID = ''' +  @ProductID + ''', '
	Set @Sql3 = @Sql3 + 'FieldName = ''' + @FieldName + ''', '  + @Sql1

	-- // Write the values to Rpt_ProductHistoryWorking
	EXEC(@Sql3)
	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = 'Error #1246|usp_RptBuildProductCompactor: Error inserting records into Rpt_ProductHistoryWorking: CallDate: ' + Convert(varchar, @CallDate, 101) + ', ClientID: ' + @ClientID + ', ProductID: ' +  @ProductID + ', FieldName: '  + @FieldName
		Return
	End

	-- // Cycle through the working table. Transfer the information to Rpt_ProductHistory only if the FieldData field has a value other than 0.
	DECLARE WorkingCursor CURSOR FOR Select EnrollerID, CallDate, ClientID, ProductID, FieldName, FieldData  From Rpt_ProductHistoryWorking Where FieldData <> 0
	Open WorkingCursor
	Fetch Next From WorkingCursor Into @EnrollerID, @CallDate, @ClientID, @ProductID, @FieldName, @FieldData
	While @@Fetch_Status = 0
	Begin
		If (@FieldData <> 0) and (@FieldName <> 'WeeklyPremium')
		Begin
			Set @Sql3 = 'Insert Into Rpt_ProductHistory (AddDate, EnrollerID, CallDate, ClientID, ProductID, FieldName, FieldData) '
			Set @Sql3 = @Sql3 + ' Values '
			Set @Sql3 = @Sql3 + '(''' +      Convert(varchar, @AddDate, 120) + ''', '''    + @EnrollerID + ''', ''' + Convert(varchar, @CallDate, 102) + ''', ''' + @ClientID + ''', ''' + @ProductID + ''', ''' + @FieldName + ''', ' + Cast(@FieldData as varchar(10)) + ')'
			Exec(@Sql3)
			SELECT @Err = @@error
			IF @Err <> 0
			Begin
				Select @RetVal = 'Error #1248|usp_RptBuildProductCompactor: Error inserting records into Rpt_ProductHistoryWorking: CallDate: ' + Convert(varchar, @CallDate, 101) + ', ClientID: ' + @ClientID + ', ProductID: ' +  @ProductID + ', FieldName: '  + @FieldName
				Return
			End

		End
	Fetch Next From WorkingCursor Into @EnrollerID, @CallDate, @ClientID, @ProductID, @FieldName, @FieldData
	End

	-- // Clean up the curso
	Close WorkingCursor
	Deallocate WorkingCursor

	Set @RetVal = '0|Success'


SET NOCOUNT OFF






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO











/*
Declare @FieldName varchar(20)
Declare @ProductID varchar(20)
Declare @ClientID varchar(20)
Declare @RetVal varchar(500)
Declare @Sql2 varchar(5000)
Declare @CallDate datetime
Set @CallDate = '3/15/2010'
Set @ClientID = 'STANTEC'
Set @ProductID = 'TMARKCOMBO'
Set @FieldName = 'EZ1_5'
SELECT @Sql2 = ''
SELECT @RetVal = '0|Success'
 exec usp_RptBuildProductSegmentExtended1  @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT, @RetVal OUTPUT
Select RetVal = @Sql2

*/



CREATE                    PROCEDURE usp_RptBuildProductSegmentExtended1(@CallDate smalldatetime, @ClientID varchar(20), @ProductID varchar(20), @FieldName varchar(20), @Sql2 varchar(5000) OUTPUT, @RetVal varchar(500) OUTPUT)
AS

SET NOCOUNT ON

BEGIN

	Declare @RelationCodeClause varchar(70)
	Declare @AddlCondition varchar(400)

	If @FieldName IN ('EE', 'EEEZ')
		Set @RelationCodeClause = 'alt.RelationCode = ''E'''
	Else If @FieldName IN ('SP', 'SPEZ')
		Set @RelationCodeClause = 'alt.RelationCode = ''S'''
	Else If @FieldName = 'EESP'
		Set @RelationCodeClause = 'alt.RelationCode = ''ESP'''
	Else If @FieldName = 'EECH'
		Set @RelationCodeClause = 'alt.RelationCode = ''ECH'''
	Else If @FieldName = 'FAM'
		Set @RelationCodeClause = 'alt.RelationCode = ''FAM'''
	Else If @FieldName = 'DEP'
		Set @RelationCodeClause = 'alt.RelationCode = ''C'''


	IF @ProductID IN ('MOTERMEXPRESS15', 'MOTERMEXPRESS20', 'MOTERMEXPRESS30', 'MOTERMCOMPLETE15', 'MOTERMCOMPLETE20', 'MOTERMCOMPLETE30')
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	ELSE IF @ProductID = 'EYEMEDVIS3'
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	ELSE IF @ProductID IN ('ALLSTATECI', 'ALLSTATESTD', 'ALLSTATETERM', 'ALLSTATEACC', 'ALLSTATEDI')
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	ELSE IF @ProductID = 'ALLSTATEUL'
	Begin
		If @FieldName IN ('EE', 'EEEZ', 'SP', 'SPEZ')
		Begin
			Set @AddlCondition = case
				when @FieldName IN ('EE', 'SP')  then @RelationCodeClause
				when @FieldName IN ('EEEZ', 'SPEZ') then @RelationCodeClause + ' AND alt.Tier = ''EZ'''
				end
				Select @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, @AddlCondition, 0, 'et2')
		End
		Else
		Begin
			exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
		End
	End


	Else If @ProductID = 'AMGENCANCER'	
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	Else If @ProductID IN ('BMCCI', 'BMDI', 'BMWLIFE')
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	Else If @ProductID = 'BONDS'
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	-- // The spreadsheet handles the COKC FSA as a single entity. The CCM product maintain section handles it as two separate entities. Using one of the definitions used by the product section enables the app to properly handle the product.
	Else If @ProductID = 'COKCFSA_MED'	
	Begin
		If @FieldName = 'MEDFSA'
			Select @Sql2 = 'MEDFSA = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.ClientSubProductID = ''COKCFSA_MED''', 0, 'et2')
		Else If @FieldName = 'DEPFSA'
			Select @Sql2 = 'DEPFSA = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.ClientSubProductID = ''COKCFSA_DEP''', 0, 'et2')
	End

/*
		IF @FieldName = 'MEDFSA'
			SELECT 'MEDFSA = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.ClientSubProductID = ''COKCFSA_MED''', 'et2')
		ELSE IF @FieldName = 'DEPFSA'
			SELECT 'DEPFSA = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.ClientSubProductID = ''COKCFSA_DEP''', 'et2')
		ELSE
			exec  usp_RptBuildProductSegmentExtended2 @CurDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
*/


	Else If @ProductID IN ('EYEMEDVIS', 'EYEMEDVIS2')
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	Else If @ProductID = 'HARTFORDLTD'
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	Else If @ProductID = 'TMARKCANCER'
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	Else If @ProductID = 'TMARKCOMBO' OR @ProductID = 'TMARKCI'
	Begin
		If @FieldName IN ('EE', 'EESP', 'EECH', 'FAM')
			Select @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, @RelationCodeClause, 0, 'et2')
		Else If @FieldName = 'EZ1_5'
			Select @Sql2 = @FieldName + ' = ' + dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.Tier = ''EZ1_5''', 0, 'et2')
		Else
			exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
	End


/*
	ELSE IF @ProductID = 'TMARKCOMBO'
		IF @FieldName = 'EE'
			SELECT 'EE = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''E'' AND (alt.Tier is null or alt.Tier = ''EZ1_5'')', 'et2')
		ELSE IF @FieldName = 'EESP'
			SELECT 'EESP = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''ESP'' AND (alt.Tier is null or alt.Tier = ''EZ1_5'')', 'et2')
		ELSE IF @FieldName = 'EECH'
			SELECT 'EECH = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''ECH'' AND (alt.Tier is null or alt.Tier = ''EZ1_5'')', 'et2')
		ELSE IF @FieldName = 'FAM'
			SELECT 'FAM = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''FAM'' AND (alt.Tier is null or alt.Tier = ''EZ1_5'')', 'et2')
		ELSE IF @FieldName = 'EZ1_5'
			SELECT 'EZ1_5 = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode IN  (''E'', ''ESP'', ''ECH'', ''FAM'')  AND alt.Tier = ''EZ1_5''', 'et2')
		ELSE
			exec  usp_RptBuildProductSegmentExtended2 @CurDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
*/


/*
	ELSE IF @ProductID = 'TMARKUL'
		IF @FieldName = 'EE'
			SELECT 'EE = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''E'' AND (alt.Tier is null or alt.Tier = ''''  or alt.Tier IN  (''EZ1_10'', ''EZ2_5'', ''EZ1_5''))', 'et2')
		ELSE IF @FieldName = 'SP'
			SELECT 'SP = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''S'' AND (alt.Tier is null or alt.Tier = '''' or alt.Tier IN  (''EZ1_10'', ''EZ2_5'', ''EZ1_5''))', 'et2')
		ELSE IF @FieldName = 'DEP'
			SELECT 'DEP = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''DEP'' AND (alt.Tier is null or alt.Tier = '''' or alt.Tier IN  (''EZ1_10'', ''EZ2_5'', ''EZ1_5''))', 'et2')
		ELSE IF @FieldName = 'EZ1_10'
			SELECT 'EZ1_10 = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.Tier = ''EZ1_10''', 'et2')
		ELSE IF @FieldName = 'EZ2_5'
			SELECT 'EZ2_5 = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.Tier = ''EZ2_5''', 'et2')
		ELSE IF @FieldName = 'EZ1_5'
			SELECT 'EZ1_5 = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.Tier = ''EZ1_5''', 'et2')
		ELSE
			exec  usp_RptBuildProductSegmentExtended2 @CurDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
*/

	Else If @ProductID = 'TMARKUL' OR @ProductID = 'TMARKUL2' OR @ProductID = 'TMARKULNY'
	Begin
		If @FieldName In ('EE', 'SP', 'DEP', 'EZ1_10', 'EZ2_5', 'EZ1_5')
		Begin
			Set @AddlCondition = case
				when @FieldName In ('EE', 'SP', 'DEP') then @RelationCodeClause
				when @FieldName = 'EZ1_10' then 'alt.Tier = ''EZ1_10'''
				when @FieldName = 'EZ2_5' then 'alt.Tier = ''EZ2_5'''
				when @FieldName = 'EZ1_5' then 'alt.Tier = ''EZ1_5'''
				end
				Select @Sql2 = @FieldName + ' = ' + dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, @AddlCondition, 0, 'et2')
		End
		Else
		Begin
			exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
		End
	End

	Else If @ProductID = 'TRANSCANCERSELPLUS'
	Begin
		If @FieldName In ('EE', 'EECH', 'FAM', 'PLAN1', 'PLAN2')
		Begin
			Set @AddlCondition = case 
				when @FieldName In  ('EE', 'EECH', 'FAM') then @RelationCodeClause
				when @FieldName = 'PLAN1' then 'alt.Tier = ''PLAN1'''
				when @FieldName = 'PLAN2' then 'alt.Tier = ''PLAN2'''
				end
				Select @Sql2 = @FieldName + ' = ' + dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, @AddlCondition, 0, 'et2')
		End
		Else
		Begin
			exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
		End
	End

/*
	Else If @ProductID = 'TRANSCANCERSELPLUS'
		IF @FieldName = 'EE'
			SELECT 'EE = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''E'' AND (alt.Tier is null or alt.Tier IN  (''PLAN1'', ''PLAN2''))', 'et2')
		ELSE IF @FieldName = 'EECH'
			SELECT 'EECH = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''ECH'' AND (alt.Tier is null or alt.Tier IN  (''PLAN1'', ''PLAN2''))', 'et2')
		ELSE IF @FieldName = 'FAM'
			SELECT 'FAM = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.RelationCode = ''FAM'' AND (alt.Tier is null or alt.Tier IN  (''PLAN1'', ''PLAN2''))', 'et2')
		ELSE IF @FieldName = 'PLAN1'
			SELECT 'PLAN1 = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.Tier= ''PLAN1''', 'et2')
		ELSE IF @FieldName = 'PLAN2'
			SELECT 'PLAN2 = ' +  dbo.ufn_BuildProductSegment3_Extended(@ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CurDate, 'alt.Tier= ''PLAN2''', 'et2')
		ELSE
			exec  usp_RptBuildProductSegmentExtended2 @CurDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
*/	
	Else If @ProductID IN ('TRANSDI', 'TRANSTERM', 'TRANSUL2')
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT

	Else If @ProductID = 'UNUMACC'
		exec  usp_RptBuildProductSegmentExtended2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT


	Select @RetVal = @ProductID
	Return
END

SET NOCOUNT OFF


























GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO














/*
Declare @FieldName varchar(20)
Declare @ProductID varchar(20)
Declare @ClientID varchar(20)
Declare @RetVal varchar(500)
Declare @Sql2 varchar(5000)
Declare @CallDate datetime
Set @CallDate = '8/3/2009'
Set @ClientID = 'COKC'
Set @ProductID = 'TRANSCANCERSELPLUS'
Set @FieldName = 'WEEKLYPREMIUM'
SELECT @Sql2 = ''
SELECT @RetVal = '0|Success'
 exec usp_RptBuildProductSegmentExtended2  @CallDate,  @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT
Select Sql2 = @Sql2
*/



CREATE     PROCEDURE usp_RptBuildProductSegmentExtended2(@CallDate smalldatetime, @ClientID varchar(20), @ProductID varchar(20), @FieldName varchar(20), @Sql2 varchar(5000) OUTPUT)
AS

SET NOCOUNT ON

BEGIN



	IF @FieldName = 'YES'
	Begin
		SELECT @Sql2 = 'Yes =  (SELECT Count (*) FROM ProjectReports..EmpTransmittal etOuter '
		SELECT @Sql2 = @Sql2 + 'INNER JOIN '
		SELECT @Sql2 = @Sql2 + '(SELECT DISTINCT etYes.ActivityID '
		SELECT @Sql2 = @Sql2 + 'FROM ProjectReports..EmpProductTransmittal eptYes '
		SELECT @Sql2 = @Sql2 + 'INNER JOIN ProjectReports..EmpTransmittal etYes ON eptYes.ActivityID = etYes.ActivityID '
		SELECT @Sql2 = @Sql2 + 'INNER JOIN BVI..Client bviclient ON etYes.ClientID = bviclient.ClientID '
		SELECT @Sql2 = @Sql2 + dbo.ufn_GetClientJoin(@ClientID, 'etYes')
		SELECT @Sql2 = @Sql2 + 'WHERE eptYes.LogicalDelete = 0 AND '
		SELECT @Sql2 = @Sql2 + dbo.ufn_GetClientWhere(@ClientID, 'etYes')
		SELECT @Sql2 = @Sql2 + 'eptYes.LicensedEnroller = u.UserID AND eptYes.ProductID = ''' + @ProductID + ''' AND '

		-- 11/27/2009
		SELECT @Sql2 = @Sql2 + 'dbo.ufn_IsAPSLegal(eptYes.ActivityID, eptYes.AppID, eptYes.ClientID, eptYes.ProductID) = 1 AND '

		SELECT @Sql2 = @Sql2 + 'dbo.ufn_IsDateEqual(etYes.CallStartTime, ''' +  Convert(varchar, @CallDate, 101) + ''') = 1 AND etYes.SupervisorApprovalDate IS NOT NULL) '
		SELECT @Sql2 = @Sql2 + 'YesSpecial ON etOuter.ActivityID = YesSpecial.ActivityID) '
	End

	Else If @FieldName = 'EE'
		Select @Sql2 = 'EE = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''E''', 0, 'et2')

	Else If @FieldName = 'EE+1'
		Select @Sql2 = 'EE+1 = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''EE+1''', 0, 'et2')

	Else If @FieldName = 'SP'
		Select @Sql2 = 'SP = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''S''', 0, 'et2')

	Else If @FieldName = 'CH'
		Select @Sql2 = 'CH = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''C''', 0, 'et2')

	Else If @FieldName = 'EESP'
		Select @Sql2 = 'EESP = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''ESP''', 0, 'et2')

	Else If @FieldName = 'EECH'
		Select @Sql2 = 'EECH = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''ECH''', 0, 'et2')

	Else If @FieldName = 'SPCH'
		Select @Sql2 = 'SPCH = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''SPC''', 0, 'et2')

	Else If @FieldName = 'FAM'
		Select @Sql2 = 'FAM = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''FAM''', 0, 'et2')

	Else If @FieldName = 'DEP'
		Select @Sql2 = 'DEP = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'Count (*)', 'et2.CallStartTime', @CallDate, 'alt.RelationCode = ''C''', 0, 'et2')

	Else If @FieldName = 'WEEKLYPREMIUM'	
		Select @Sql2 = 'WEEKLYPREMIUM = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'IsNull(Convert(varchar, SUM(ept.WeeklyPremium)), 0)', 'et2.CallStartTime', @CallDate, '', 0, 'et2')

	Else If @FieldName = 'ANNUALPREMIUM'
		Select @Sql2 = 'ANNUALPREMIUM = ' +  dbo.ufn_RptBuildProductSegment3(1, @ClientID, @ProductID, 'IsNull(Convert(varchar, ROUND(SUM(ISNULL(ept.WeeklyPremium, 0)) * 52, 0)), 0)', 'et2.CallStartTime', @CallDate, '', 0, 'et2')
END

SET NOCOUNT OFF

















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--sp_helptext usp_RptBuildProductSegmentStandard2

  /*  
Declare @CallDate datetime  
Declare @RetVal varchar(200)  
Declare @Sql2 varchar(5000)  
Declare @ClientID varchar(20)  
Declare @ProductID varchar(20)  
Declare @FieldName varchar(20)  
Set @ClientID = 'Stantec'  
Set @ProductID = 'TMarkCombo'  
Set @CallDate = '12/02/2009'  
  
Set @FieldName = 'EZ1_5'  
Set @FieldName = 'EESP'  
--SELECT @Sql2 = 'bbbbb'  
 exec usp_RptBuildProductSegmentStandard1  @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 output,   @RetVal output  
Select Sql2 = @Sql2  
print @Sql2  
*/  
  
  
CREATE           PROCEDURE usp_RptBuildProductSegmentStandard1(@CallDate smalldatetime, @ClientID varchar(20), @ProductID varchar(20), @FieldName varchar(20), @Sql2 varchar(5000) OUTPUT, @RetVal varchar(200) OUTPUT)  
AS  
  
SET NOCOUNT ON  
  
BEGIN  
  
 Declare @RelationCodeClause varchar(70)  
 Declare @AddlCondition varchar(400)  
  
 If @FieldName IN ('EE', 'EEEZ')  
  Set @RelationCodeClause = ' (aps.RelationCode IN (''E'', '''') OR aps.RelationCode is null) '  
 Else If @FieldName IN ('SP', 'SPEZ')  
  Set @RelationCodeClause = ' aps.RelationCode = ''S'' '  
 Else If @FieldName = 'EESP'  
  Set @RelationCodeClause = ' aps.RelationCode In (''ES'', ''ESP'') '  
 Else If @FieldName = 'EECH'  
  Set @RelationCodeClause = ' aps.RelationCode In (''EC'', ''ECH'') '  
 --Else If @FieldName = 'FAM'  
  -- No action. Legal and TMarkCombo are inconsistent.  
 --Else If @FieldName = 'DEP'  
  -- No action. TransUL and TMarkUL are inconsistent.  
 --Else If @FieldName IN ('WEEKLYPREMIUM', 'ANNUALPREMIUM', 'YES', 'EZ1_5', 'EZ1_10', 'EZ2_5')  
  -- No action.  
  
  
 -- // AIG  
 If @ProductID IN ('AIGACC', 'AIGCCI')  
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  
 -- // AllState  
 Else IF @ProductID In ('ALLSTATECCI','ALLSTATECANCER', 'ALLSTATECANCER2', 'ALLSTATECI', 'ALLSTATEACC', 'ALLSTATESTD', 'ALLSTATETERM', 'ALLSTATECAN', 'ALLSTATECAN2', 'ALLSTATEDI')  
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  
 Else If @ProductID = 'ALLSTATEUL'  
 Begin  
  If @FieldName IN ('EE', 'SP', 'EEEZ', 'SPEZ')  
  Begin  
   Set @AddlCondition = case  
    when @FieldName IN ('EE', 'SP') then '(apd.FieldName IS NULL OR apd.FieldName = ''Future Purchase Option Rider'') '  
    when @FieldName IN ('EEEZ', 'SPEZ') then 'apd.FieldName = ''Future Purchase Option Rider'' '  
    end  
   Set @AddlCondition = @AddlCondition + ' AND ' + @RelationCodeClause  
   Set @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, @AddlCondition, 1, 'et2')  
  End  
  If @FieldName In ('YES', 'DEP', 'WEEKLYPREMIUM', 'ANNUALPREMIUM')  
  Begin  
   exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  End  
 End  
  
 -- // American General  
 Else If @ProductID = 'AMGENCANCER'  
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  

  
 -- // Bonds
 Else If @ProductID IN ('BONDS')   
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  
  
 -- // Boston Mutual  
 Else If @ProductID IN ('BMCCI', 'BMWLIFE', 'BMDI')  
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  
  
 -- // EyeMed  
 Else If @ProductID IN ('EYEMEDVIS', 'EYEMEDVIS3')  
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  
  
 -- // Both Legal Plan of America and Hyatt Legal use the same ProductID -- Legal -- in IAMS.  
 Else If @ProductID IN ('LEGAL')  
 Begin  
  If @FieldName = 'YES'  
   exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  Else If @FieldName = 'FAM'  
  Begin  
   Set @AddlCondition = 'aps.RelationCode = ''E'' '  
   Set @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, @AddlCondition, 0, 'et2')  
  End  
  Else If @FieldName IN ('WEEKLYPREMIUM', 'ANNUALPREMIUM')  
   exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  Else  
   exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
 End  
  
  
 -- // Trustmark  
 Else If @ProductID = 'TMARKACC'  
 Begin  
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
 End  
 Else If @ProductID in ('TMARKCOMBO','TMARKCANCER','TMARKCI')
 Begin  
  If @FieldName IN ('EE', 'EESP', 'EECH', 'FAM')  
  Begin  
   --Set @AddlCondition = '(apd.FieldName IS NULL Or apd.FieldName = ''Riders'' AND CharIndex(''EZV'', apd.FieldData) > 0) AND '  
   --Set @AddlCondition = @AddlCondition + @RelationCodeClause  
   Set @AddlCondition =  @RelationCodeClause  
   Set @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, 'TMARKCOMBO', 'Count (*)', 'aps.AppDate', @CallDate, @AddlCondition, 0, 'et2')  
  End  
  
  Else If @FieldName = 'EZ1_5'  
  Begin  
   --Set @AddlCondition = '(apd.FieldName = ''Riders'' AND CharIndex(''EZV'', apd.FieldData) > 0) '  
   Set @AddlCondition = '(apd.FieldName = ''EZValue'' AND apd.FieldData = ''true'')'  
   Set @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, 'TMARKCOMBO', 'Count (*)', 'aps.AppDate', @CallDate, @AddlCondition, 1, 'et2')  
  End  
  Else  
   exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
 End  
  
  
 Else If @ProductID in ('TMARKUL','TMARKUL2','TMARKULNY')
 Begin  
  IF @FieldName IN ('EE', 'SP')  
  Begin  
   Set @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, 'TMARKUL', 'Count (*)', 'aps.AppDate', @CallDate, @RelationCodeClause, 0, 'et2')  
  End  
  Else If @FieldName IN ('EZ1_5', 'EZ1_10', 'EZ2_5')  
  Begin  
   Set @AddlCondition = case  
    when @FieldName = 'EZ1_5' then 'apd.FieldName = ''Riders'' AND CharIndex(''EZV1'', apd.FieldData) > 0'  
    when @FieldName = 'EZ1_10' then 'apd.FieldName = ''Riders'' AND CharIndex(''EZV2'', apd.FieldData) > 0'  
    when @FieldName = 'EZ2_5' then 'apd.FieldName = ''Riders'' AND CharIndex(''EZV3'', apd.FieldData) > 0'  
    end  
   Set @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, 'TMARKUL', 'Count (*)', 'aps.AppDate', @CallDate, @AddlCondition, 1, 'et2')  
  End  
  Else If @FieldName = 'DEP'  
  Begin  
   Set @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, 'aps.RelationCode = ''C''', 0, 'et2')  
  End  
  Else  
  Begin  
   exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT   
  End  
 End  
  
 -- // TransAmerica  
 Else If @ProductID IN ('TRANSACC', 'TRANSCCI', 'TRANSDI', 'TRANSTERM','TRANSCANCERSELPLUS')  
 Begin  
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
 End  
  
 Else If @ProductID in ('TRANSUL','TRANSUL2')
 Begin  
  If @FieldName IN ('EE', 'SP', 'EEEZ', 'SPEZ', 'DEP')  
  Begin  
   Set @AddlCondition = case  
    --when @FieldName IN ('EE', 'SP') then  '(apd.FieldName IS NULL OR apd.FieldData = ''True'')'  
  
    --when @FieldName = 'EE' then  '(apd.FieldName IS NULL OR apd.FieldData = ''True'')  AND aps.RelationCode = ''E'''  
    --when @FieldName = 'SP' then  '(apd.FieldName IS NULL OR apd.FieldData = ''True'')  AND aps.RelationCode = ''S'''  
    --when @FieldName IN ('EEEZ', 'SPEZ') then  'apd.FieldName = ''EZValue'' AND apd.FieldData = ''True'''  
    --when @FieldName = 'DEP' then  '(apd.FieldName IS NULL OR apd.FieldData = ''True'')  AND aps.RelationCode IN (''C'',''D'')'  
  
    when @FieldName = 'EE' then  'aps.RelationCode = ''E'''  
    when @FieldName = 'EEEZ' then  'apd.FieldName  = ''True'' AND aps.RelationCode = ''E'''  
    when @FieldName = 'SP' then  'aps.RelationCode = ''S'''  
    when @FieldName = 'SPEZ' then  'apd.FieldName  = ''True'' AND aps.RelationCode = ''S'''  
    when @FieldName = 'DEP' then  'aps.RelationCode IN (''C'',''D'')'  
    end  
  
   Set @Sql2 = @FieldName + ' = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, 'TRANSUL', 'Count (*)', 'aps.AppDate', @CallDate, @AddlCondition, 1, 'et2')  
  End  
  If @FieldName IN ('YES', 'WEEKLYPREMIUM', 'ANNUALPREMIUM')  
  Begin  
   exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  End  
 End  
  
  
  
 -- // Unum  
 Else If @ProductID = 'UNUMACC'  
  exec  usp_RptBuildProductSegmentStandard2 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT  
  
  
 Select @RetVal = @ProductID  
 Return  
  
END  
  
SET NOCOUNT OFF  
  
  
  


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
















/*

Declare @SqlTest varchar(1000)
Declare @ThisDate datetime
set @ThisDate = '2/28/2011'
set @SqlTest = ''

 exec usp_RptBuildProductSegmentStandard2   @ThisDate,  'C3', 'TRANSCCI', 'AnnualPremium', @SqlTest 

*/



CREATE                  PROCEDURE usp_RptBuildProductSegmentStandard2(@CallDate smalldatetime, @ClientID varchar(20), @ProductID varchar(20), @FieldName varchar(20), @Sql2 varchar(5000) OUTPUT)
AS

SET NOCOUNT ON

BEGIN

	IF @FieldName = 'YES'
	Begin
		Set @Sql2 = 'YES = (SELECT Count (*) FROM EmpTransmittal etOuterYes '
		--Set @Sql2 = 'YES = (SELECT Count (*) FROM EmpProductTransmittal eptOuterYes '
		Set @Sql2 = @Sql2 + 'INNER JOIN '
		Set @Sql2 = @Sql2 + '(SELECT DISTINCT etInner.ActivityID '
		Set @Sql2 = @Sql2 + 'FROM ProjectReports..EmpTransmittal etInner '
		Set @Sql2 = @Sql2 + 'INNER JOIN ProjectReports..EmpProductTransmittal eptInner ON etInner.ActivityID = eptInner.ActivityID  '

		-- // 1/25/2011
		--Set @Sql2 = @Sql2 + 'INNER JOIN IAMS..AppsAndPolsSummary apsInner ON etInner.ActivityID = apsInner.ActivityID '
		Set @Sql2 = @Sql2 + 'INNER JOIN IAMS..AppsAndPolsSummary apsInner ON eptInner.AppID = apsInner.AppID '


		Set @Sql2 = @Sql2 + 'INNER JOIN BVI..Client bviclient ON apsInner.ClientID = bviclient.ClientID '
		Set @Sql2 = @Sql2 +  dbo.ufn_GetClientJoin(@ClientID, 'apsInner')
		Set @Sql2 = @Sql2 + 'WHERE eptInner.LogicalDelete = 0 AND '
		Set @Sql2 = @Sql2 + 'etInner.SupervisorApprovalDate IS NOT NULL AND '
		Set @Sql2 = @Sql2 + dbo.ufn_GetClientWhere(@ClientID, 'eptInner')
		Set @Sql2 = @Sql2 + 'eptInner.LicensedEnroller = u.UserID AND '
		Set @Sql2 = @Sql2 + 'eptInner.ProductID = ''' + @ProductID + ''' AND '
		Set @Sql2 = @Sql2 + 'apsInner.BVIAppStatus <> ''CANCELLED'' AND '
		Set @Sql2 = @Sql2 + 'dbo.ufn_IsDateEqual(etINNER.CallStartTime, ''' + Convert(varchar, @CallDate, 101)  + ''') = 1 AND '
		Set @Sql2 = @Sql2  + 'dbo.ufn_IsTestID(bviclient.ClientID, etINNER.EmpID) = 0) '
		Set @Sql2 = @Sql2 + 'YesSpecial ON etOuterYes.ActivityID = YesSpecial.ActivityID) '
	End

	IF @FieldName = 'EE'
		Select @Sql2 = 'EE = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, '(aps.RelationCode = ''E'' or aps.RelationCode = '''' or aps.RelationCode is null)', 0, 'et2')

	ELSE IF @FieldName = 'SP'
		Select @Sql2 = 'SP = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, 'aps.RelationCode = ''S''', 0, 'et2')

	ELSE IF @FieldName = 'CH'
		Select @Sql2 = 'CH = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, 'aps.RelationCode = ''C''', 0, 'et2')

	ELSE IF @FieldName = 'EESP'
		Select @Sql2 = 'EESP = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, 'aps.RelationCode IN (''ES'', ''ESP'')', 0, 'et2')

	ELSE IF @FieldName = 'EECH'
		Select @Sql2 = 'EECH = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, 'aps.RelationCode IN (''EC'', ''ECH'')', 0, 'et2')

	ELSE IF @FieldName = 'SPCH'
		Select @Sql2 = 'SPCH = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, 'aps.RelationCode = ''SPC''', 0, 'et2')

	ELSE IF @FieldName = 'FAM'
		Select @Sql2 = 'FAM = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, 'aps.RelationCode IN (''FAM'', ''F'')', 0, 'et2')

	ELSE IF @FieldName = 'DEP'
		Select @Sql2 = 'DEP = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'Count (*)', 'aps.AppDate', @CallDate, 'aps.RelationCode = ''D''', 0, 'et2')

	ELSE IF @FieldName = 'WEEKLYPREMIUM'
		Select @Sql2 = 'WEEKLYPREMIUM = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'IsNull(Convert(varchar, SUM(ept.WeeklyPremium)), 0)', 'aps.AppDate', @CallDate, '', 0, 'et2')

	ELSE IF @FieldName = 'ANNUALPREMIUM'
		Select @Sql2 = 'ANNUALPREMIUM = ' +  dbo.ufn_RptBuildProductSegment3(0, @ClientID, @ProductID, 'IsNull(Convert(varchar, ROUND(SUM(ISNULL(ept.WeeklyPremium, 0)) * 52, 0)), 0)', 'aps.AppDate', @CallDate, '', 0, 'et2')

END

SET NOCOUNT OFF



















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO












/*
Select count (*) from rpt_callhistory
select count (*) from rpt_producthistory
delete rpt_callhistory
Delete rpt_producthistory

Declare @RetVal varchar(500)
Declare @CallDate smalldatetime
Declare @AddDate datetime
Declare @ClientID varchar(20)
Set @CallDate = '9/23/2009'
Set @ClientID = 'BureauVeritas'
Set @AddDate = '9/23/2009'
Set @RetVal = '0|Success'
exec usp_RptBuildProductSegments @CallDate, @AddDate, @ClientID, @RetVal  OUTPUT
Select RetVal = @RetVal
*/


CREATE       PROCEDURE usp_RptBuildProductSegments(@CallDate smalldatetime, @AddDate datetime,  @ClientID varchar(20), @RetVal varchar(500) OUTPUT)
AS

SET NOCOUNT ON

	-- // Declarations
	Declare @Err int
	Declare @UseStandardInd bit

	Declare @Sql2 varchar(5000)
	Declare @Sql3 varchar(2000)


	Declare @ProductTable Table(RecID int IDENTITY, ProductID varchar(20))
	Declare @ProductTableRecordcount int
	Declare @ProductList varchar(200)
	Declare @ProductTableRecID int
	Declare @ProductID varchar(20)

	Declare @FieldTable TABLE(RecID int, FieldName varchar(20))
	Declare @FieldTableRecordcount int	
	Declare @FieldName varchar(20)
	Declare @FieldTableRecID int

	-- // ProductTable
	--INSERT INTO @ProductTable SELECT SegmentID FROM Excel_ClusterSegment WHERE ClusterID = @ClientID

	INSERT INTO @ProductTable SELECT ccp.BVIProductID 
	FROM Excel_ClusterSegment ecs
	INNER JOIN Config_ClientProduct ccp on ecs.SegmentID = ccp.ClientProductID
 	WHERE ClusterID = @ClientID

	SELECT @Err = @@error
	IF @Err <> 0
	Begin
		Select @RetVal = '1240|usp_RptBuildProductSegment: Error inserting records into @ProductTable'
		Return
	End
	SELECT @ProductTableRecordcount = Count (*) FROM @ProductTable


	-- // Delete the records that are about to be replaced
	DELETE Rpt_ProductHistory WHERE ClientID = @ClientID AND dbo.ufn_IsDateEqual(CallDate, @CallDate) = 1


	-- // ProductTable loop
	Set @ProductTableRecID = 1
	While @ProductTableRecID <= @ProductTableRecordcount
	Begin

		Select @ProductID = ProductID FROM @ProductTable WHERE RecID = @ProductTableRecID
--If @ProductID <> 'ZZZZZZZZ'
--Begin

--print @ProductID

			-------------------------------------------------------------------------------------------------------------------------
	             		-- // Use IAMS or Alt for this product
			Set  @UseStandardInd = 0
			IF(SELECT Count (*) FROM ClientProduct_Extended WHERE  ClientID =  @ClientID  AND ClientProductID = @ProductID AND dbo.ufn_IsDateBetween(@CallDate, StartDate, EndDate) = 1) = 0
			Begin
				SELECT @UseStandardInd = 1
			End
			-------------------------------------------------------------------------------------------------------------------------


			-------------------------------------------------------------------------------------------------------------------------
			DELETE @FieldTable
			INSERT INTO @FieldTable Select * FROM dbo.ufn_GetTableFromList(@ProductID)
			Set @FieldTableRecID = 1
			Select @FieldTableRecordcount =  Count (*) FROM @FieldTable
			While @FieldTableRecID <= @FieldTableRecordcount
			Begin
				Select @FieldName = FieldName FROM @FieldTable WHERE RecID = @FieldTableRecID

--If @FieldName <> 'ZZZZZZZZ'
--Begin
					-------------------------------------------------------------------------------------------------------------------------
					SELECT @Sql2 = ''
					If @UseStandardInd = 1
					Begin
						exec usp_RptBuildProductSegmentStandard1 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT, @RetVal OUTPUT
						If @RetVal <> '0|Success'
						Begin
							Return
						End
					End
					Else
					Begin
						exec usp_RptBuildProductSegmentExtended1 @CallDate, @ClientID, @ProductID, @FieldName, @Sql2 OUTPUT, @RetVal OUTPUT
						If @RetVal <> '0|Success'
						Begin
							Return
						End
					End

					exec usp_RptBuildProductCompactor @CallDate, @AddDate,  @ClientID, @ProductID, @FieldName, @Sql2, @RetVal OUTPUT
					If @RetVal <> '0|Success'
					Begin
						Return
					End
	
					--Select  Replace(@Sql2, '''', '''''')

					-------------------------------------------------------------------------------------------------------------------------
--End
				Set @FieldTableRecID = @FieldTableRecID + 1

			End
			-------------------------------------------------------------------------------------------------------------------------
--End
		Set @ProductTableRecID = @ProductTableRecID + 1

	End


SET NOCOUNT OFF

















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




--     usp_RptBuildProductSegmentsThisClient '8/20/2009', 'Morgans'



CREATE     PROCEDURE usp_RptBuildProductSegmentsThisClient(@CurDate smalldatetime, @ClientID varchar(20))
AS

SET NOCOUNT ON


BEGIN

	-- // Sql
	Declare @Sql1 varchar(8000)
	Declare @Sql2 varchar(5000)

	-- // Product
	Declare @ProductID varchar(20)
	Declare @ClientProductExtendedMember int
	Declare @UseIAMSInd bit
	Declare @UseAltInd bit

	-- // Column
	Declare @ColumnTable TABLE(RecID int, ColumnName varchar(20))
	Declare @ColumnRecID int
	Declare @ColumnName varchar(20)
	Declare @ColumnList varchar(200)
	Declare @FldList varchar(200)
	Declare @SplitOn varchar(1)
	Declare @ColumnRecordcount int
	Declare @ColumnIterator int

	-- // Build a cursor for the products associated with this client
	DECLARE ProductCursor CURSOR FOR SELECT SegmentID FROM Excel_ClusterSegment WHERE ClusterID = @ClientID
	OPEN ProductCursor
	FETCH NEXT FROM ProductCursor into @ProductID
	WHILE @@FETCH_STATUS = 0
	Begin

             		-- // Use IAMS or Alt for this product
		SELECT  @UseIAMSInd = 0
		SELECT  @UseAltInd = 0

		IF(SELECT Count (*) FROM ClientProduct_Extended WHERE  ClientID =  @ClientID  AND ClientProductID = @ProductID AND dbo.ufn_IsDateBetween(@CurDate, StartDate, EndDate) = 1) = 0
		Begin
			SELECT @UseIAMSInd = 1
		End
		ELSE
		Begin
			SELECT @UseAltInd = 1
		End


		-- // Start the Sql
		SELECT @Sql1 = 'SELECT '
		SELECT @Sql1 = @Sql1 +  'EnrollerID = u.UserID,       '
		SELECT @Sql1 = @Sql1 + 'ClientID = ''' + @ClientID + ''',         '
		SELECT @Sql1 = @Sql1 + 'EnrollerDay = ''' + Cast(@CurDate as varchar(20)) + ''',          '

		-- // ColumnList, FldList
		SELECT @ColumnList = Columns FROM Excel_SegmentConfigure WHERE SegmentId = @ProductID
		--SELECT @FldList = '(EnrollerID, ClientID, EnrollerDay, ' + Replace(@ColumnList, '|', ',') + ') '
		SELECT @FldList = Replace(@ColumnList, '|', ',') + ') '

		-- // ColumnTable 
		Select @ColumnRecID = 1
		Delete @ColumnTable
		Set @SplitOn = '|'
		While (CharIndex(@SplitOn, @ColumnList) > 0)
		Begin
			Insert Into @ColumnTable (RecID, ColumnName)  Select RecID = @ColumnRecID, ColumnName = ltrim(rtrim(Substring(@ColumnList, 1, CharIndex(@SplitOn, @ColumnList) -1) ) )
			Set @ColumnList = Substring(@ColumnList, Charindex(@SplitOn, @ColumnList)+1, len(@ColumnList))
			Select @ColumnRecID = @ColumnRecID + 1
		End
		INSERT INTO @ColumnTable (RecID, ColumnName) Select RecID = @ColumnRecID, ColumnName = ltrim(rtrim(@ColumnList))

		-- // Iterate through the products for this client
		SELECT @ColumnIterator = 1
		SELECT @ColumnRecordcount = Count (*) FROM @ColumnTable		
		WHILE @ColumnIterator <= @ColumnRecordcount
		Begin
			SELECT @ColumnName = ColumnName FROM @ColumnTable WHERE RecID = @ColumnIterator

			IF @UseIAMSInd = 1
			Begin

				SELECT @Sql2 = ''
				exec usp_RptBuildProductSegmentStandard1 @CurDate, @ClientID, @ProductID, @ColumnName, @Sql2 OUTPUT
				SELECT @Sql1 = @Sql1 + @Sql2


			End
			If @UseAltInd = 1
			Begin
				SELECT @Sql2 = ''
				exec usp_RptBuildProductSegmentExtended1 @CurDate, @ClientID, @ProductID, @ColumnName, @Sql2 OUTPUT
				SELECT @Sql1 = @Sql1 + @Sql2
			End
			
			--Select Results = Cast(@CurDate as varchar(20))  + '      ' + @ClientID + '    '  + @ProductID  + '   ' + @ColumnName
			--Select Results = Cast(@CurDate as varchar(20))  + '      ' + @ClientID + '    '  + Cast(@ColumnRecordcount as varchar(20))  + '   ' + Cast(@ColumnIterator as varchar(20)) + '    '  + @ProductID  + '   ' + @ColumnName
			SELECT @ColumnIterator = @ColumnIterator + 1
		End


	-- // Complete the Sql
	SELECT @Sql1 = @Sql1 + 'FROM UserManagement..Users u '
	SELECT @Sql1 = @Sql1 + 'WHERE '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EmpTransmittal etMain '
	SELECT @Sql1 = @Sql1 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
	SELECT @Sql1 = @Sql1 + 'WHERE etMain.LogicalDelete = 0 AND etMain.EnrollerID = u.UserID AND '
	SELECT @Sql1 = @Sql1 +  'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''' +  Cast(@CurDate as varchar(20)) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND etMain.SupervisorApprovalDate IS NOT NULL AND '
	SELECT @Sql1 = @Sql1  + 'dbo.ufn_IsTestID(bviclient.ClientID, etMain.EmpID) = 0) > 0 '
	SELECT @Sql1 = @Sql1  + 'OR '
	SELECT @Sql1 = @Sql1 + '(SELECT Count(*) FROM ProjectReports..EnrollerDate edMain '
	SELECT @Sql1 = @Sql1 + 'WHERE edMain.EnrollerID = u.UserID AND dbo.ufn_IsDateEqual(edMain.ProjectDate, ''' + Cast(@CurDate as varchar(20)) + ''') = 1 '
	SELECT @Sql1 = @Sql1 + 'AND edMain.SupervisorApproval = 1) >0 '


	-- // Delete the records that are about to be replaced
	--DELETE Rpt_CallHistory WHERE ClientID = @ClientID AND dbo.ufn_IsDateEqual(EnrollerDay, @CurDate) = 1

	-- // Insert the records into Rpt_CallHistory
	--SELECT @Sql1 = 'INSERT INTO Rpt_CallHistory ' + @Fldlist + @Sql1
	 --exec(@Sql1)





		FETCH NEXT FROM ProductCursor into @ProductID
	End

	Close ProductCursor
	Deallocate ProductCursor

END


SET NOCOUNT OFF




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






/*
Select count (*) from rpt_callhistory
select count (*) from rpt_producthistory

Delete rpt_callhistory
delete rpt_producthistory
Declare @RetVal varchar(500)
Set @RetVal = 'nono'
exec usp_RptBuildStart
Select RetVal = @RetVal
*/


CREATE    PROCEDURE  usp_RptBuildStart

AS


-- // How does the update work?
-- // 1. Set the initial date: 8/3/2009
-- // 2. Get a list of call dates from EmpTransmittal and EmpProductTransmittal whose change date falls after the last AddDate in the Rpt_CallHistory and Rpt_CallProduct tables
-- // 3. Delete all of the records with this call date from Rpt_CallHistory and Rpt_ProductHistory
-- // 4. Update all of the records with these call dates in Rpt_CallHistory and Rpt_ProductHistory




SET NOCOUNT ON
	Declare @RetVal varchar(500)
	Declare @AddDate datetime
	Declare @LastAddDate datetime
	Declare @Rpt_CallHistoryRecordcount int

	-- // Client table
	Declare @ClientID varchar(20)
	Declare @Client Table(RecID int IDENTITY,ClientID varchar(20))
	Declare @ClientRecID int
	Declare @ClientRecordcount int

	SELECT @RetVal = '0|Success'
	SELECT @AddDate = DateAdd(hour, -5, GetUTCDate())
	SELECT @Rpt_CallHistoryRecordcount = Count (*) FROM Rpt_CallHistory

	------------------------------------------------------------------------------------------------------------------------------------------------
	-- // CallDate table

	Declare @FirstCallDate smalldatetime
	Declare @CallDateRecID int
	Declare @CallDate smalldatetime
	Declare @CallDateRaw Table(CallDate smalldatetime)
	Declare @CallDateTable Table(RecID int IDENTITY, CallDate smalldatetime)
	Declare @CallDateRecordcount int

	Set @FirstCallDate = '8/3/2009'

	SELECT @Rpt_CallHistoryRecordcount = Count (*) FROM Rpt_CallHistory
	If @Rpt_CallHistoryRecordcount = 0
	Begin
		SELECT @LastAddDate = '8/3/2009'
	End
	Else
	Begin
		SELECT @LastAddDate = Max(AddDate) FROM Rpt_CallHistory
	End

	-- // Get a list of call dates from EmpTransmittal and EmpProductTransmittal whose ChangeDate falls after the AddDate in the Rpt_CallHistory and Rpt_CallProduct tables
	INSERT INTO @CallDateRaw 
	SELECT Convert(varchar, et.CallStartTime, 101) FROM EmpTransmittal et
	INNER JOIN EmpProductTransmittal ept on et.ActivityID = ept.ActivityID
	WHERE (et.CallStartTime >= @FirstCallDate) AND 
	(et.ChangeDate > @LastAddDate OR ept.ChangeDate > @LastAddDate)

	-- // Clean up by eliminating duplicate.
	INSERT INTO @CallDateTable SELECT DISTINCT CallDate FROM @CallDateRaw ORDER BY CallDate
	Select @CallDateRecordcount = Count (*) FROM @CallDateTable

	------------------------------------------------------------------------------------------------------------------------------------------------

	-- // Build the Client table
	INSERT INTO @Client SELECT SegmentID FROM Excel_SegmentConfigure WHERE SegmentType = 'CLIENT'
	Select @ClientRecordcount = Count (*) FROM @Client

	------------------------------------------------------------------------------------------------------------------------------------------------
	-- // Start: Loop through the dates and clients. Call usp_RptBuildClientSegment1 and usp_RptBuildProductSegment
	-- // for each combination of date and client.

	SELECT @CallDateRecID = 1
	WHILE @CallDateRecID <= @CallDateRecordcount
	Begin	

		SELECT @CallDate = CallDate FROM @CallDateTable WHERE RecID = @CallDateRecID

print @CallDate

		Select @ClientRecID = 1
		WHILE @ClientRecID <= @ClientRecordcount
		Begin
			SELECT @ClientID = ClientID FROM @Client WHERE RecID = @ClientRecID
			--Select Results =  Cast(@CallData as varchar(12)) + '    ' +  @ClientID  + Cast(@ClientRecID as varchar(2))
print @ClientID
			-- // Build client segment this date and client
			exec usp_RptBuildClientSegment1 @CallDate, @AddDate, @ClientID, @RetVal OUTPUT
			If @RetVal <> '0|Success'
			Begin
				Select RetVal = @RetVal
				Return
			End

			--  // Build product segments this date and client
			exec usp_RptBuildProductSegments @CallDate, @AddDate, @ClientID, @RetVal OUTPUT
			If @RetVal <> '0|Success'
			Begin
				Select RetVal = @RetVal
				Return
			End

			Set @ClientRecID = @ClientRecID + 1
		End

		SELECT @CallDateRecID = @CallDateRecID + 1

	End

	Select RetVal = @RetVal
	Return


SET NOCOUNT ON




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




--usp_BuildTempDatesUpdateRequired

CREATE    PROCEDURE usp_RptBuildTempDatesUpdateRequired
AS

BEGIN

	Declare @DateTable Table(UpdateDate smalldatetime)


	-- // Drop temp table if is exists. Create new table.
	IF OBJECT_ID ('ProjectReports.dbo.TempDatesUpdateRequired','U') IS NOT NULL 
	Begin
		DROP Table TempDatesUpdateRequired
	End
	CREATE Table TempDatesUpdateRequired(RedID int IDENTITY, UpdateDate smalldatetime)

	INSERT INTO @DateTable
	SELECT DISTINCT NeedsUpdate = Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 
	FROM EmpTransmittal 
	ORDER  BY  Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 

	INSERT INTO TempDatesUpdateRequired(UpdateDate) SELECT * FROM @DateTable	


	Select * from TempDatesUpdateRequired

END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




--        usp_RptStartMemOnly

CREATE     Procedure usp_RptStartMemOnly

AS
SET NOCOUNT ON

	-- // DATE TABLE
	Declare @CurDate smalldatetime
	Declare @DateUpdate Table(RecID int IDENTITY, DateUpdate smalldatetime)
	Declare @DateUpdateIterator int
	Declare @DateUpdateRecordcount int
	INSERT INTO @DateUpdate
	SELECT DISTINCT NeedsUpdate = Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 
	FROM EmpTransmittal 
	Set @DateUpdateIterator = 1
	Select @DateUpdateRecordcount = Count (*) FROM @DateUpdate


	-- // CLIENT TABLE
	Declare @CurClientID varchar(20)
	Declare @Client Table(RecID int IDENTITY,ClientID varchar(20))
	Declare @ClientIterator int
	Declare @ClientRecordcount int
	INSERT INTO @Client SELECT SegmentID FROM Excel_SegmentConfigure WHERE SegmentType = 'CLIENT'
	Set @ClientIterator = 1
	Select @ClientRecordcount = Count (*) FROM @Client


	-- // Iterate through the date and client tables
	WHILE @DateUpdateIterator <= @DateUpdateRecordcount
	Begin	
		SELECT @CurDate = DateUpdate FROM @DateUpdate WHERE RecID = @DateUpdateIterator
		Set @DateUpdateIterator = @DateUpdateIterator + 1
		
		Set @ClientIterator = 1
		WHILE @ClientIterator <= @ClientRecordcount
		Begin
			SELECT @CurClientID = ClientID FROM @Client WHERE RecID = @ClientIterator
			--Select Results =  Cast(@CurDate as varchar(12)) + '    ' +  @CurClientID
			exec usp_RptBuildClientSegment1 @CurDate,  @CurClientID
			Set @ClientIterator = @ClientIterator + 1
		End

	End


SET NOCOUNT ON





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




-- usp_RptUpdateData '123'

CREATE    Procedure usp_RptUpdateData
(
  @Tags VARCHAR(2000)
)
AS
SET NOCOUNT ON

Declare @CurRecID int
DECLARE @Recordcount int
DECLARE @tagIds Table (RecID int IDENTITY, UserID varchar(20))
Insert into @tagIds Select top 10 UserID from usermanagement..users

Select @Recordcount = Count (*) FROM @tagIds
Set @CurRecID = 1
WHILE @CurRecID <= @Recordcount
Begin

	Declare @Space int
	Select * FROM @tagIds WHERE RecID = @CurRecID
	Set @CurRecID = @CurRecID + 1
End


Select * from  @tagIds

SET NOCOUNT ON






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_Rpt_All]              
(              
 @Date1 datetime = '1/1/1800',                  
 @Date2 datetime ='1/1/1800',            
 @EnrollerID varchar(100)=NULL,            
 @ClientID varchar(100)=NULL,      
 @ProductID varchar(100)=NULL,      
 @Location varchar(50)=NULL
)              
AS              
----------------------------------------------------------            
--Daterange defaults to yesterday            
--EnrollerID defaults to "All"            
--ClientID defauts to "All"       
--CallingSource is 1 if CCM, 0 or null if the DTS procedure     
----------------------------------------------------------              
            
BEGIN              
              
   DECLARE @Yesterday datetime              
   DECLARE @Today datetime              
   DECLARE @StartDate datetime      
   DECLARE @EndDate datetime      
               
   SELECT @Yesterday = CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())-1) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))            
   SELECT @Today =  CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))              
            
   SELECT @StartDate = CASE WHEN ISNULL(@Date1,'1/1/1800') = '1/1/1800' THEN @Yesterday ELSE @Date1 END            
   SELECT @EndDate = CASE WHEN ISNULL(@Date2,'1/1/1800') = '1/1/1800' THEN @Today ELSE @Date2 END            
      
      
Select     
 EnrollerID=CASE WHEN et.EnrollerID is NULL Then 'ALL' ELSE et.EnrollerID END,    
 PremiumPerManday=CASE WHEN et.EnrollerID is NULL THEN 'N/A' ELSE CAST(CAST((ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0) / (MAX(edp.TotalHours) / 8.0) as decimal(20,2)) as varchar(100)) END,    
 PremiumPerInterview=CAST(ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 / (COUNT(*) * 1.0) as decimal(20,2)),    
 CloseRatio=CONVERT(varchar(10),CONVERT(decimal(5,1),ROUND((SUM(ISNULL(IAMS.Enrollments,0)) * 1.0) / (COUNT(*) * 1.0) * 100,1)))+'%',    
 Interviews=COUNT(*),    
 Enrollments=SUM(ISNULL(IAMS.Enrollments,0)),    
 TotalPremium=CAST(ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 as decimal(20,2)),    
 ManHours=CASE WHEN et.EnrollerID is NULL THEN 'N/A' ELSE CAST(MAX(edp.TotalHours) as varchar(100)) END,    
 --EnrollmentTime=MAX(edp.EnrollHours),    
 EnrPct=CASE WHEN et.EnrollerID is NULL THEN 'N/A' ELSE CONVERT(varchar(10),CONVERT(decimal(5,1),ROUND((MAX(edp.EnrollHours) * 1.0) / (MAX(edp.TotalHours) * 1.0) * 100,1))) + '%' END,    
 StartDate=CONVERT(varchar(50),@StartDate,101),    
 EndDate=CONVERT(varchar(50),@EndDate,101),    
 ClientID=CASE WHEN @ClientID is NULL THEN 'ALL' ELSE @ClientID END,    
 ProductID=CASE WHEN @ProductID is NULL THEN 'ALL' ELSE @ProductID END,    
 Location=ISNULL((Select locationID from UserManagement..Users where UserID=et.EnrollerID),ISNULL(@Location,'ALL'))    
From UserManagement..Users as U    
INNER JOIN ProjectReports..EmpTransmittal as et    
ON    
 U.UserID = Et.EnrollerID    
 AND u.ROLE in ('ENROLLER','SUPERVISOR')     
 AND u.CompanyID='BVI'    
 AND (@Location=u.LocationID or @Location is NULL)    
 AND et.CallStartTime BETWEEN @StartDate  AND @EndDate    
 AND et.CallEndTime is not NULL       
 AND et.ActivityTypeCode = 'CALL'    
 AND (@EnrollerID = et.EnrollerID or @EnrollerID is NULL)    
 AND (@ClientID = et.ClientID or @ClientID is NULL)    
INNER JOIN     
 (Select     
  EnrollerID,     
  TotalHours=SUM(TotalHours),    
  EnrollHours=SUM(EnrollHours)    
 from ProjectReports..EnrollerDateProject     
 Where    
  (ClientID = @ClientID or @ClientID is NULL)    
  and (EnrollerID = @EnrollerID or @EnrollerID is NULL)    
  and ProjectDate BETWEEN @StartDate and @EndDate
  and ProjectDate<>@EndDate 
-- and ProjectDate >= @StartDate 
 -- and ProjectDate < @EndDate    
 Group by EnrollerID) as edp    
ON    
 et.EnrollerID = edp.EnrollerID    
LEFT OUTER JOIN     
 (SELECT ActivityID, MonthlyPremium=SUM(MonthlyPremium), Enrollments=1    
 FROM    
 IAMS..Log_AppsAndPolsSummary     
 WHERE    
  BVIAppStatus in ('NEW','POLICY')    
  and ChangeTypeCode='INSERT'    
  and (@ProductID = ProductID or @ProductID is NULL)  
  and AppID in (    
   Select AppID from IAMS..Log_AppsAndPolsSummary   
  Where ChangeTypeCode in ('INSERT','DELETE')   
   group by AppID HAVING count(*) = 1)  
 Group by ActivityID) as IAMS    
ON    
 IAMS.ActivityID=et.ActivityID    
Group By et.EnrollerID with ROLLUP    
Order by CASE WHEN et.EnrollerID is NULL THEN 100000000 ELSE CAST((ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0) / (MAX(edp.TotalHours) / 8.0) as decimal(20,2)) END desc    
      
END      


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_Rpt_All_102011]          
(          
 @Date1 datetime = '1/1/1800',              
 @Date2 datetime ='1/1/1800',        
 @EnrollerID varchar(100)=NULL,        
 @ClientID varchar(100)=NULL,  
 @ProductID varchar(100)=NULL,  
 @Location varchar(50)=NULL                 
)          
AS          
----------------------------------------------------------        
--Daterange defaults to yesterday        
--EnrollerID defaults to "All"        
--ClientID defauts to "All"        
----------------------------------------------------------          
        
BEGIN          
          
   DECLARE @Yesterday datetime          
   DECLARE @Today datetime          
   DECLARE @StartDate datetime  
   DECLARE @EndDate datetime  
           
   SELECT @Yesterday = CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())-1) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))        
   SELECT @Today =  CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))          
        
   SELECT @StartDate = CASE WHEN ISNULL(@Date1,'1/1/1800') = '1/1/1800' THEN @Yesterday ELSE @Date1 END        
   SELECT @EndDate = CASE WHEN ISNULL(@Date2,'1/1/1800') = '1/1/1800' THEN @Today ELSE @Date2 END        
  
  
SELECT  
      EnrollerID,  
      Location=(Select locationID from UserManagement..Users where UserID=EnrollerID),  
      ManHours=SUM(DATEDIFF(mi,et.CallStartTime,et.CallEndTime)/60.0),  
      Interviews=COUNT(*),  
      AnnualPremium=CAST(ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 as money),  
      Enrollments=SUM(ISNULL(IAMS.Enrollments,0)),  
      PremiumPerManday=CAST(  
 CASE SUM(DATEDIFF(mi,et.CallStartTime,et.CallEndTime))   
 WHEN 0 THEN 0 ELSE  
  ISNULL(SUM(IAMS.MonthlyPremium),0) * 12 /  
  SUM(DATEDIFF(mi,et.CallStartTime,et.CallEndTime)/60.0) / 8.0 END as money),  
      CloseRatio=(SUM(ISNULL(IAMS.Enrollments,0)) * 1.0) / (COUNT(*) * 1.0),  
      PremiumPerInterview=CAST(ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 /(COUNT(*) * 1.0) as money),  
      StartDate=@StartDate,  
      EndDate=@EndDate,  
      ClientID=CASE WHEN @ClientID is NULL THEN 'ALL' ELSE @ClientID END,  
      ProductID=CASE WHEN @ProductID is NULL THEN 'ALL' ELSE @ProductID END  
From UserManagement..Users as U  
INNER JOIN ProjectReports..EmpTransmittal as et  
ON  
      U.UserID = Et.UserID  
      AND u.ROLE in ('ENROLLER','SUPERVISOR')  
      AND u.CompanyID='BVI'  
      AND (@Location=u.LocationID or @Location is NULL)  
      AND et.CallStartTime BETWEEN  @StartDate AND @EndDate  
      AND et.CallEndTime is NOT NULL    
      AND et.ActivityTypeCode = 'CALL'  
      AND (@EnrollerID = et.EnrollerID or @EnrollerID is NULL)  
      AND (@ClientID = et.ClientID or @ClientID is NULL)  
LEFT OUTER JOIN  
      (SELECT ActivityID, MonthlyPremium=SUM(MonthlyPremium), Enrollments=1  
      FROM  
      IAMS..Log_AppsAndPolsSummary  
      WHERE  
            BVIAppStatus in ('NEW','POLICY')  
            and ChangeTypeCode='INSERT'  
            and (@ProductID = ProductID or @ProductID is NULL  
                  and AppID in (  
                        Select AppID from IAMS..Log_AppsAndPolsSummary Where  
   ChangeTypeCode in ('INSERT','DELETE') group by AppID HAVING count(*) = 1))  
      Group by ActivityID) as IAMS  
ON  
      IAMS.ActivityID=et.ActivityID  
Group By EnrollerID  WITH Rollup
Order by EnrollerID  
  
END  


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_Rpt_All_NameOnly]              
(              
 @Date1 datetime = '1/1/1800',                  
 @Date2 datetime ='1/1/1800',            
 @EnrollerID varchar(100)=NULL,            
 @ClientID varchar(100)=NULL,      
 @ProductID varchar(100)=NULL,      
 @Location varchar(50)=NULL
)              
AS              
----------------------------------------------------------            
--Daterange defaults to yesterday            
--EnrollerID defaults to "All"            
--ClientID defauts to "All"        
----------------------------------------------------------              
            
BEGIN              
              
   DECLARE @Yesterday datetime              
   DECLARE @Today datetime              
   DECLARE @StartDate datetime      
   DECLARE @EndDate datetime      
               
   SELECT @Yesterday = CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())-1) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))            
   SELECT @Today =  CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))              
            
   SELECT @StartDate = CASE WHEN ISNULL(@Date1,'1/1/1800') = '1/1/1800' THEN @Yesterday ELSE @Date1 END            
   SELECT @EndDate = CASE WHEN ISNULL(@Date2,'1/1/1800') = '1/1/1800' THEN @Today ELSE @Date2 END            
      
      
Select     
 EnrollerID=CASE WHEN et.EnrollerID is NULL Then 'ALL' ELSE et.EnrollerID END,    
 U.Lastname + ', ' + U.Firstname as EnrollerName
From UserManagement..Users as U    
INNER JOIN ProjectReports..EmpTransmittal as et    
ON    
 U.UserID = Et.EnrollerID    
 AND u.ROLE in ('ENROLLER','SUPERVISOR')     
 AND u.CompanyID='BVI'    
 AND (@Location=u.LocationID or @Location is NULL)    
 AND et.CallStartTime BETWEEN @StartDate  AND @EndDate    
 AND et.CallEndTime is not NULL       
 AND et.ActivityTypeCode = 'CALL'    
 AND (@EnrollerID = et.EnrollerID or @EnrollerID is NULL)    
 AND (@ClientID = et.ClientID or @ClientID is NULL)    
INNER JOIN     
 (Select     
  EnrollerID,     
  TotalHours=SUM(TotalHours),    
  EnrollHours=SUM(EnrollHours)    
 from ProjectReports..EnrollerDateProject     
 Where    
  (ClientID = @ClientID or @ClientID is NULL)    
  and (EnrollerID = @EnrollerID or @EnrollerID is NULL)    
  and ProjectDate BETWEEN @StartDate and @EndDate
  and ProjectDate<>@EndDate 
 Group by EnrollerID) as edp    
ON    
 et.EnrollerID = edp.EnrollerID    
LEFT OUTER JOIN     
 (SELECT ActivityID, MonthlyPremium=SUM(MonthlyPremium), Enrollments=1    
 FROM    
 IAMS..Log_AppsAndPolsSummary     
 WHERE    
  BVIAppStatus in ('NEW','POLICY')    
  and ChangeTypeCode='INSERT'    
  and (@ProductID = ProductID or @ProductID is NULL)  
  and AppID in (    
   Select AppID from IAMS..Log_AppsAndPolsSummary   
  Where ChangeTypeCode in ('INSERT','DELETE')   
   group by AppID HAVING count(*) = 1)  
 Group by ActivityID) as IAMS    
ON    
 IAMS.ActivityID=et.ActivityID    
WHERE U.Lastname + ', ' + U.Firstname IS NOT NULL
Group By et.EnrollerID, U.Lastname + ', ' + U.Firstname --with ROLLUP    
Order By U.Lastname + ', ' + U.Firstname      
END      
    
  

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_Rpt_All_NoManHours]                
(                
 @Date1 datetime = '1/1/1800',                    
 @Date2 datetime ='1/1/1800',              
 @EnrollerID varchar(100)=NULL,              
 @ClientID varchar(100)=NULL,        
 @ProductID varchar(100)=NULL,        
 @Location varchar(50)=NULL  
)                
AS                
----------------------------------------------------------              
--Daterange defaults to yesterday              
--EnrollerID defaults to "All"              
--ClientID defauts to "All"         
--For Midday use when enrollers haven't updated
--EnrollerDateProject
----------------------------------------------------------                
              
BEGIN                
                
   DECLARE @Yesterday datetime                
   DECLARE @Today datetime                
   DECLARE @StartDate datetime        
   DECLARE @EndDate datetime        
                 
   SELECT @Yesterday = CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())-1) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))              
   SELECT @Today =  CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))                
              
   SELECT @StartDate = CASE WHEN ISNULL(@Date1,'1/1/1800') = '1/1/1800' THEN @Yesterday ELSE @Date1 END              
   SELECT @EndDate = CASE WHEN ISNULL(@Date2,'1/1/1800') = '1/1/1800' THEN @Today ELSE @Date2 END              
        
        
Select       
 EnrollerID=CASE WHEN et.EnrollerID is NULL Then 'ALL' ELSE et.EnrollerID END,      
 PremiumPerManday='N/A',
 PremiumPerInterview=CAST(ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 / (COUNT(*) * 1.0) as decimal(20,2)),      
 CloseRatio=CONVERT(varchar(10),CONVERT(decimal(5,1),ROUND((SUM(ISNULL(IAMS.Enrollments,0)) * 1.0) / (COUNT(*) * 1.0) * 100,1)))+'%',      
 Interviews=COUNT(*),      
 Enrollments=SUM(ISNULL(IAMS.Enrollments,0)),      
 TotalPremium=CAST(ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 as decimal(20,2)),      
 ManHours='N/A',      
 --EnrollmentTime=MAX(edp.EnrollHours),      
 EnrPct='N/A',
 StartDate=CONVERT(varchar(50),@StartDate,101),      
 EndDate=CONVERT(varchar(50),@EndDate,101),      
 ClientID=CASE WHEN @ClientID is NULL THEN 'ALL' ELSE @ClientID END,      
 ProductID=CASE WHEN @ProductID is NULL THEN 'ALL' ELSE @ProductID END,      
 Location=ISNULL((Select locationID from UserManagement..Users where UserID=et.EnrollerID),ISNULL(@Location,'ALL'))      
From UserManagement..Users as U      
INNER JOIN ProjectReports..EmpTransmittal as et      
ON      
 U.UserID = Et.EnrollerID      
 AND u.ROLE in ('ENROLLER','SUPERVISOR')       
 AND u.CompanyID='BVI'      
 AND (@Location=u.LocationID or @Location is NULL)      
 AND et.CallStartTime BETWEEN @StartDate  AND @EndDate      
 AND et.CallEndTime is not NULL         
 AND et.ActivityTypeCode = 'CALL'      
 AND (@EnrollerID = et.EnrollerID or @EnrollerID is NULL)      
 AND (@ClientID = et.ClientID or @ClientID is NULL)         
LEFT OUTER JOIN       
 (SELECT ActivityID, MonthlyPremium=SUM(MonthlyPremium), Enrollments=1      
 FROM      
 IAMS..Log_AppsAndPolsSummary       
 WHERE      
  BVIAppStatus in ('NEW','POLICY')      
  and ChangeTypeCode='INSERT'      
  and (@ProductID = ProductID or @ProductID is NULL)    
  and AppID in (      
   Select AppID from IAMS..Log_AppsAndPolsSummary     
  Where ChangeTypeCode in ('INSERT','DELETE')     
   group by AppID HAVING count(*) = 1)    
 Group by ActivityID) as IAMS      
ON      
 IAMS.ActivityID=et.ActivityID      
Group By et.EnrollerID with ROLLUP      
Order by CASE WHEN et.EnrollerID is NULL THEN 100000000 ELSE CAST((ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 / 8.0) as decimal(20,2)) END desc      
        
END        
  


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE
PROCEDURE [dbo].[usp_Rpt_All_Old]          
(          
 @Date1 datetime = '1/1/1800',              
 @Date2 datetime ='1/1/1800',        
 @EnrollerID varchar(100)=NULL,        
 @ClientID varchar(100)=NULL,  
 @ProductID varchar(100)=NULL,  
 @Location varchar(50)=NULL                 
)          
AS          
----------------------------------------------------------        
--Daterange defaults to yesterday        
--EnrollerID defaults to "All"        
--ClientID defauts to "All"        
----------------------------------------------------------          
        
BEGIN          
          
   DECLARE @Yesterday datetime          
   DECLARE @Today datetime          
   DECLARE @StartDate datetime  
   DECLARE @EndDate datetime  
           
   SELECT @Yesterday = CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())-1) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))        
   SELECT @Today =  CONVERT(varchar(2),datepart(mm,getdate())) + '/' + CONVERT(varchar(2),datepart(dd,getdate())) + '/' + CONVERT(varchar(4),datepart(yy,getdate()))          
        
   SELECT @StartDate = CASE WHEN ISNULL(@Date1,'1/1/1800') = '1/1/1800' THEN @Yesterday ELSE @Date1 END        
   SELECT @EndDate = CASE WHEN ISNULL(@Date2,'1/1/1800') = '1/1/1800' THEN @Today ELSE @Date2 END        
  
  
SELECT  
      EnrollerID,  
      Location=(Select locationID from UserManagement..Users where UserID=EnrollerID),  
      ManHours=SUM(DATEDIFF(mi,et.CallStartTime,et.CallEndTime)/60.0),  
      Interviews=COUNT(*),  
      AnnualPremium=CAST(ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 as money),  
      Enrollments=SUM(ISNULL(IAMS.Enrollments,0)),  
      PremiumPerManday=CAST(  
 CASE SUM(DATEDIFF(mi,et.CallStartTime,et.CallEndTime))   
 WHEN 0 THEN 0 ELSE  
  ISNULL(SUM(IAMS.MonthlyPremium),0) * 12 /  
  SUM(DATEDIFF(mi,et.CallStartTime,et.CallEndTime)/60.0) / 8.0 END as money),  
      CloseRatio=(SUM(ISNULL(IAMS.Enrollments,0)) * 1.0) / (COUNT(*) * 1.0),  
      PremiumPerInterview=CAST(ISNULL(SUM(IAMS.MonthlyPremium),0) * 12.0 /(COUNT(*) * 1.0) as money),  
      StartDate=@StartDate,  
      EndDate=@EndDate,  
      ClientID=CASE WHEN @ClientID is NULL THEN 'ALL' ELSE @ClientID END,  
      ProductID=CASE WHEN @ProductID is NULL THEN 'ALL' ELSE @ProductID END  
From UserManagement..Users as U  
INNER JOIN ProjectReports..EmpTransmittal as et  
ON  
      U.UserID = Et.UserID  
      AND u.ROLE in ('ENROLLER','SUPERVISOR')  
      AND u.CompanyID='BVI'  
      AND (@Location=u.LocationID or @Location is NULL)  
      AND et.CallStartTime BETWEEN  @StartDate AND @EndDate  
      AND et.CallEndTime is NOT NULL    
      AND et.ActivityTypeCode = 'CALL'  
      AND (@EnrollerID = et.EnrollerID or @EnrollerID is NULL)  
      AND (@ClientID = et.ClientID or @ClientID is NULL)  
LEFT OUTER JOIN  
      (SELECT ActivityID, MonthlyPremium=SUM(MonthlyPremium), Enrollments=1  
      FROM  
      IAMS..Log_AppsAndPolsSummary  
      WHERE  
            BVIAppStatus in ('NEW','POLICY')  
            and ChangeTypeCode='INSERT'  
            and (@ProductID = ProductID or @ProductID is NULL  
                  and AppID in (  
                        Select AppID from IAMS..Log_AppsAndPolsSummary Where  
   ChangeTypeCode in ('INSERT','DELETE') group by AppID HAVING count(*) = 1))  
      Group by ActivityID) as IAMS  
ON  
      IAMS.ActivityID=et.ActivityID  
Group By EnrollerID  
Order by EnrollerID  
  
END  


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE usp_Rpt_EnrollerProductivity(@FirstReportDate smalldatetime, @LastReportDate datetime)
AS

SET NOCOUNT ON

SELECT DISTINCT rph.EnrollerID,
EnrollerName = u.LastName + ', ' + u.FirstName,

Interviews = (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1),

Enrolled = (SELECT IsNull(Sum(Enrolled), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1),

EnrollHours = (SELECT IsNull(Sum(EnrollHours), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1),

AdminHours = (SELECT IsNull(Sum(AdminHours), 0) +  IsNull(Sum(TrainHours), 0) +  IsNull(Sum(CoachHours), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1),

TotalHours = (SELECT IsNull(Sum(AdminHours), 0) + IsNull(Sum(EnrollHours), 0) +  IsNull(Sum(TrainHours), 0) +  IsNull(Sum(CoachHours), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1),

Premium = 
	(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0)
	FROM Rpt_ProductHistory
	WHERE EnrollerID = rph.EnrollerID AND
	dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1 AND
	FieldName = 'AnnualPremium'),

PremiumMD = case
	when (SELECT IsNull(Sum(EnrollHours), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1) = 0 then 0
	else
		(
		(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0)
		FROM Rpt_ProductHistory
		WHERE EnrollerID = rph.EnrollerID AND
		dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1 AND
		FieldName = 'AnnualPremium')
		/
		((SELECT Sum(EnrollHours) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1) / 8)
		)
	end,

PremiumInterviewed =  case
	when (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1) = 0 then 0
	else
		(
		(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0)
		FROM Rpt_ProductHistory
		WHERE EnrollerID = rph.EnrollerID AND
		dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1 AND
		FieldName = 'AnnualPremium')
		/
		(SELECT Sum(Interviewed) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1)
		)
	end,

PremiumEnrolled = case
	when (SELECT IsNull(Sum(Enrolled), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1) = 0 then 0
	else
		(
		(Select IsNull(SUM(Cast(IsNull(FieldData, 0) as decimal(10,2))), 0)
		FROM Rpt_ProductHistory
		WHERE EnrollerID = rph.EnrollerID AND
		dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1 AND
		FieldName = 'AnnualPremium')
		/
		(SELECT Sum(Enrolled) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1)
		)
	end,

Ratio = case
	when (SELECT IsNull(Sum(Interviewed), 0) FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1) = 0 then 0
	else
		(
		(SELECT IsNull(Sum(Enrolled), 0) * 1.0 FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1)
		/
		(SELECT IsNull(Sum(Interviewed), 0) * 1.0 FROM Rpt_CallHistory WHERE EnrollerID = rph.EnrollerID AND  dbo.ufn_IsDateBetween(CallDate, @FirstReportDate, @LastReportDate) = 1)
		)
	end










FROM Rpt_ProductHistory rph
LEFT JOIN UserManagement..Users u ON rph.EnrollerID = u.UserID
WHERE dbo.ufn_IsDateBetween(rph.CallDate, @FirstReportDate, @LastReportDate) = 1
ORDER BY EnrollerName

SET NOCOUNT OFF
















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--
--          exec usp_TempBobErrorHandling @IsGood = 0

--  usp_TempBobThurs1

CREATE  Procedure usp_TempBobErrorHandling(@IsGood bit)
AS
BEGIN

If @IsGood = 0
Begin
	Select ErrorNumber = 1, ErrorMessage = 'Interference by dogs'
End
Else
Begin
	Select ErrorNumber = 0, ErrorMessage = ''
	Select top 10 * from emptransmittal
End


END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE usp_TempCursor
AS

BEGIN

		DECLARE @ClientID varchar(20)

		-- // Get a list of lists. Set up cursor.
		DECLARE Excel_ClusterSegmentCursor CURSOR FOR SELECT DISTINCT ClusterID FROM Excel_ClusterSegment ORDER BY ClusterID
		OPEN Excel_ClusterSegmentCursor
		FETCH ClientID into @ClientID

		WHILE @@Fetch_Status = 0
		BEGIN

			EXEC usp_UpdateClientSegment1 @ClientID
		END

		CLOSE Excel_ClusterSegmentCursor
		DEALLOCATE Excel_ClusterSegmentCursor

		Declare @Dog int

END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE usp_TempCursorCallSP2
AS

BEGIN

		DECLARE @ClientID varchar(20)

		-- // Get a list of lists. Set up cursor.
		DECLARE Excel_ClusterSegmentCursor CURSOR FOR SELECT DISTINCT ClusterID FROM Excel_ClusterSegment ORDER BY ClusterID
		OPEN Excel_ClusterSegmentCursor
		FETCH ClientID into @ClientID

		WHILE @@Fetch_Status = 0
		BEGIN
Declare @Holder int
			--EXEC usp_UpdateClientSegment1 @ClientID
		END

		CLOSE Excel_ClusterSegmentCursor
		DEALLOCATE Excel_ClusterSegmentCursor

		Declare @Dog int

END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  Procedure usp_TempDBExist(@DBName varchar(20), @EmpID varchar(20))

AS

BEGIN

Declare @Sql varchar(1000)
Select @Sql = 'Declare @TableExists int '
Select @Sql = @Sql + 'Declare @IsTestID int '
Select @Sql = @Sql + 'Select @TableExists = 0 '
Select @Sql = @Sql + 'Select @IsTestID = 0 '
Select @Sql = @Sql + 'Declare @Results int '
Select @Sql = @Sql + 'Select @Results = 0 '

Select @Sql = @Sql + 'IF OBJECT_ID (''' + @DBName + '.dbo.TestIDs'',''U'') IS NOT NULL '
Select @Sql = @Sql + 'AND '
Select @Sql = @Sql + '(SELECT Count (*) FROM ' + @DBName + '.dbo.TestIDs WHERE EmpID = ''' + @EmpID + ''') =1 '
Select @Sql = @Sql + 'Select @Results = 1 ' 
Select @Sql = @Sql + 'Select Results = @Results'

exec(@Sql)
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    PROCEDURE usp_TempGetTblDef  (@tablename varchar(50))
as 
-- declare @tablename varchar(50)
-- SET @tablename = 'ImportFileEmployee'


select cast(column_name as varchar(50)) name,cast(data_type as varchar(15)) type,
length = case when isnull(character_maximum_length,0)>0 then
       character_maximum_length
   else
       numeric_precision
   end
      ,is_nullable as 'NULLS?', Cast(Column_Default as Varchar(100)) as Defaults
from information_schema.COLUMNS where table_name=@tablename

--	select * from information_schema.COLUMNS where table_name='EmpActivityLOG'









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



--BobTemp2 '123'

CREATE   Procedure usp_TempIterateMemTableGood
(
  @Tags VARCHAR(2000)
)
AS
SET NOCOUNT ON

Declare @CurRecID int
DECLARE @Recordcount int
DECLARE @tagIds Table (RecID int IDENTITY, UserID varchar(20))
Insert into @tagIds Select top 10 UserID from usermanagement..users

Select @Recordcount = Count (*) FROM @tagIds
Set @CurRecID = 1
WHILE @CurRecID <= @Recordcount
Begin

	Declare @Space int
	Select * FROM @tagIds WHERE RecID = @CurRecID
	Set @CurRecID = @CurRecID + 1
End


Select * from  @tagIds

SET NOCOUNT ON




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


--
--          exec usp_TempBobErrorHandling @IsGood = 0

--  usp_TempBobThurs1

CREATE   Procedure usp_TempReturnTableToSP1
AS
BEGIN

	--Declare 
	Declare @RetVal  varchar(20) 



	  exec  usp_TempReturnTableToSP2 @RetVal OUTPUT
Select Dog = @RetVal


END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



--  usp_TempReturnTableToSP2


--  usp_TempBobThurs1

CREATE    Procedure usp_TempReturnTableToSP2 (@RetVal varchar(20) OUTPUT)
AS
BEGIN


	--Select ErrorNumber = 1, ErrorMessage = 'Have a nice day'
Select @RetVal = '1|Have a nice day'


END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



--usp_RptStart

CREATE   Procedure usp_TempRptStartPhysicalTable

AS
SET NOCOUNT ON

	Declare @CurDate smalldatetime
	Declare @CurClientID varchar(20)

	-- // DATE TABLE
	Declare @DateUpdate Table(DateUpdate smalldatetime)
	Declare @DateUpdateIterator int
	Declare @DateUpdateRecordcount int

	-- // Build the memory table
	INSERT INTO @DateUpdate
	SELECT DISTINCT NeedsUpdate = Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 
	FROM EmpTransmittal 
	ORDER  BY  Cast(DatePart(year, ChangeDate)as varchar(4)) + '-' + dbo.ufn_PadLeft(Cast(DatePart(month, ChangeDate) as varchar(2)), 2, '0') + '-' + dbo.ufn_PadLeft(Cast(DatePart(day, ChangeDate) as varchar(2)), 2, '0') 

	-- // Create the physical table
	IF OBJECT_ID ('ProjectReports.dbo.TempDateUpdate','U') IS NOT NULL 
	Begin
		DROP Table TempDateUpdate
	End
	CREATE Table TempDateUpdate(RecID int IDENTITY, DateUpdate smalldatetime)

	-- // Copy the memory table data to the physical table
	INSERT INTO TempDateUpdate SELECT * FROM @DateUpdate

	-- // Initialize variables
	Set @DateUpdateIterator = 1
	Select @DateUpdateRecordcount = Count (*) FROM @DateUpdate
	

	-- // CLIENT TABLE
	Declare @Client Table(ClientID varchar(20))
	Declare @ClientIterator int
	Declare @ClientRecordcount int

	-- // Build the memory table
	INSERT INTO @Client  SELECT DISTINCT ClusterID FROM Excel_ClusterSegment ORDER  BY ClusterID 

	-- // Create the physical table
	IF OBJECT_ID ('ProjectReports.dbo.TempClient','U') IS NOT NULL 
	Begin
		DROP Table TempClient
	End
	CREATE Table TempClient(RecID int IDENTITY, ClientID varchar(20))

	-- // Copy the memory table data to the physical table
	INSERT INTO TempClient SELECT * FROM @Client

	-- // Initialize variables
	Set @ClientIterator = 1
	Select @ClientRecordcount = Count (*) FROM @Client

	-- // Iterate through the date and client tables
	WHILE @DateUpdateIterator <= @DateUpdateRecordcount
	Begin	
		SELECT @CurDate = DateUpdate FROM TempDateUpdate WHERE RecID = @DateUpdateIterator
		Set @DateUpdateIterator = @DateUpdateIterator + 1
		
		Set @ClientIterator = 1
		WHILE @ClientIterator <= @ClientRecordcount
		Begin
			SELECT @CurClientID = ClientID FROM TempClient WHERE RecID = @ClientIterator
			--Print Cast(@CurDate as varchar(12)) + '    ' +  @CurClientID
			Set @ClientIterator = @ClientIterator + 1
		End

	End

	-- // Clean up
	IF OBJECT_ID ('ProjectReports.dbo.TempDateUpdate','U') IS NOT NULL 
	Begin
		DROP Table TempDateUpdate
	End
	IF OBJECT_ID ('ProjectReports.dbo.TempClient','U') IS NOT NULL 
	Begin
		DROP Table TempClient
	End

SET NOCOUNT ON




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  Procedure dbo.usp_TempSortTable
As
BEGIN

	Declare @TableName varchar(25)
	Select @TableName = 'EmpTransmittal'
	
	--CREATE TABLE #TempTable
	
	SELECT * INTO #TempTable 
	FROM EmpTransmittal
	ORDER BY ClientID
	
	--SELECT * FROM #TempTable
	--DROP TABLE  #TempTable



END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






--exec usp_UpdateClientSegment1 'Armstrong', '123'

--    usp_TempSplitWithCursor 'Morgans'

CREATE PROCEDURE usp_TempSplitWithCursor(@ClientID varchar(20))
AS

BEGIN

	-- // Declarations
	Declare @ColumnName varchar(20)
	Declare @ColumnList varchar(200)
	--DECLARE @Sql1 varchar(5000)
	--DECLARE @Sql2 varchar(5000)
	DECLARE @Sql3 varchar(5000)
	Declare @SplitOn varchar(1)
	Declare @ColumnTable TABLE(ColumnName varchar(20))
	Declare @Cnt int

	-- // Get the columns field
	SELECT @ColumnList = Columns FROM Excel_SegmentConfigure WHERE SegmentId = @ClientID AND SegmentType = 'CLIENT'


	-- // Build a table with the segment elements
	Set @SplitOn = '|'
	Set @Cnt = 1
	While (CharIndex(@SplitOn, @ColumnList) > 0)
	Begin
		Insert Into @ColumnTable (ColumnName)  Select ColumnName = ltrim(rtrim(Substring(@ColumnList, 1, CharIndex(@SplitOn, @ColumnList) -1) ) )
		Set @ColumnList = Substring(@ColumnList, Charindex(@SplitOn, @ColumnList)+1, len(@ColumnList))
		Set @Cnt = @Cnt + 1
	End
	Insert Into @ColumnTable (ColumnName) Select ColumnName = ltrim(rtrim(@ColumnList))

	-- // Cycle through the elements table
	DECLARE @Element varchar(10)
	DECLARE ColumnTableCursor CURSOR FOR SELECT * FROM @ColumnTable WHERE ColumnName <> @ClientID

	OPEN ColumnTableCursor
	Fetch  ColumnTableCursor into @ClientID
	While @@Fetch_Status = 0
	Begin
		Fetch  ColumnTableCursor into @Element
print @Element

	End
	Close ColumnTableCursor
	Deallocate ColumnTableCursor




SELECT @Sql3 = 'SELECT u.UserID '
SELECT @Sql3 = @Sql3 + 'FROM UserManagement..Users u '
SELECT @Sql3 = @Sql3 + 'WHERE (SELECT Count(*) FROM ProjectReports..EmpTransmittal etMain '
SELECT @Sql3 = @Sql3 + 'INNER JOIN BVI..Client bviclient on etMain.ClientID = bviclient.ClientID '
SELECT @Sql3 = @Sql3 + 'WHERE etMain.LogicalDelete = 0 AND etMain.EnrollerID = u.UserID AND '
SELECT @Sql3 = @Sql3 + 'dbo.ufn_IsDateEqual(etMain.CallStartTime, ''08/17/2009'') = 1 AND '
SELECT @Sql3 = @Sql3 + 'u.LocationID IN (''HBG'', ''OKC'') AND '
SELECT @Sql3 = @Sql3 + 'etMain.SupervisorApprovalDate IS NOT NULL) >0 '
SELECT @Sql3 = @Sql3 + 'OR '
SELECT @Sql3 = @Sql3 + '(SELECT Count(*) FROM ProjectReports..EnrollerDate edMain '
SELECT @Sql3 = @Sql3 + 'WHERE edMain.EnrollerID = u.UserID AND '
SELECT @Sql3 = @Sql3 + 'dbo.ufn_IsDateEqual(edMain.ProjectDate, ''08/17/2009'') = 1 AND '
SELECT @Sql3 = @Sql3 + 'u.LocationID IN (''HBG'', ''OKC'') AND '
SELECT @Sql3 = @Sql3 + 'edMain.SupervisorApproval = 1) >0 '
SELECT @Sql3 = @Sql3 + 'ORDER BY u.LastName + u.FirstName'
--exec(@Sql3)

END









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


Attribute VB_Name = "GV"
Public filename As String
Public filepath As String
Public FocusBColor As ColorConstants
Public LostBColor As ColorConstants
Public ToolVersion As String
Public Tool As String
Public ImageFolder As String
Public DSN As String
Public indexer As String
Public imageLocation As String
Public imagepath As String
Public currentImage As Integer
Public data_list As String
Public trainFlag As Boolean
Public fullscreen As Boolean
Public frameWidth As Long
Public frameHeight As Long
Public imgWidth As Long
Public imgHeight As Long
Public masterDSN As String
Public tool_purpose As String
Public start_interval_time As Date
Public pcs_boxnumber As String
Public task_id As String
Public lastdoc As Boolean
Public totalImages As Long
Public start_imageindexed As String
Public last_imageindexed As String
Public count_indexed As Long
Public version As String



'*** New added ***'
Public blank_page_threshold As Long
Public backend_path As String
Public CompressionType As String
Public resolutionX As Long
Public resolutionY As Long
Public errorlog_path As String
Public destpath As String
Public package_path As String
Public from_view_next As Integer
Public OpenImage As Boolean
Public job As String
Public start_date As String
Public end_date As String
Public start_time As String
Public end_time As String
Public pre_page As Boolean
Public open_beyond_image As Integer
Public pre_boxnum As String
Public pre_boxpart As String
Public box_table_name As String
Public boxpath As String
Public indexcount As Long
Public failcount As Long
Public input_imagecount As Long
Public boxfilename As String
Public end_flag As Integer
Public current_index_page As Integer
Public preview_flag As Integer
Public lastindexpage As String
Public firstpage_list As String
Public current_firstpage As Integer
Public prev_lastname As String
Public prev_firstname As String
Public prev_TYC As String
Public prev_MM As String

'*************************'
' PCS variable            '
'*************************'
Public pcsurl As String
Public uname As String
Public pword As String
Public projectName As String
Public projectid As String
Public userid As String
Public boxid As String
Public tool_phase As String
Public skipped As Boolean
Public next_state As String
Public curr_phase As String
Public fail_count As Long
Public pass_count As Long
Public boxbarcode As String
Public boxnumber As String
Public boxpart As String
Public partnumber_length As Long
Public boxbarcode_length As Long

'sampling'
Public total_sample_field As Long
Public sample_field As Long
Public fail_rate As Long
Public indexed_field As Long
Public finish_sample As Integer
Public fail_reason As String

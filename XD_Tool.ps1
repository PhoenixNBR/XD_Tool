    ###################################################
    #                                                 #
    #    XD Tool                                      #  
    #    This script centralizes farms information    #
    #    Author : Ramzi Mahdaoui                      # 
    #    Last update 09/27/2024                       #
    #                                                 #
    ###################################################

    try
    {
	    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms
	    [System.Reflection.Assembly]::LoadFrom("Configuration\assembly\MaterialDesignColors.dll") | Out-Null
	    [System.Reflection.Assembly]::LoadFrom("Configuration\assembly\MaterialDesignThemes.Wpf.dll") | Out-Null
	    [System.Reflection.Assembly]::LoadFrom("Configuration\assembly\MaterialDesignMessageBox.dll") | Out-Null
	    asnp Citrix*
	    $ConfigPath = "$env:LOCALAPPDATA\XD_Tool\Configuration"
	    if (!(Test-Path "$env:LOCALAPPDATA\XD_Tool")) { New-Item -Path "$env:LOCALAPPDATA\XD_Tool" -Type Directory }
	    if (!(Test-Path "$ConfigPath")) { New-Item -Path "$ConfigPath" -Type Directory }
	    if (!(Test-Path "$ConfigPath\Exports")) { New-Item -Path "$ConfigPath\Exports" -Type Directory }
	    $ConfigFile = "$ConfigPath\config.xml"
	    $global:Action = $null
	    $DefaultLogo = ".\Configuration\Pictures\XD_Tool.png"
	    $global:LogoFile = "$ConfigPath\Logo.png"
	    $global:SelectLogoFile = "$ConfigPath\logo\blank.png"
	    $Green = ".\Configuration\Pictures\Green.png"
	    $Red = ".\Configuration\Pictures\Red.png"
	    $Grey = ".\Configuration\Pictures\Grey.png"
	    $pattern = '[^a-zA-Z0-9\s\.\-_]'
	    $date = get-date -Format MM_dd_yyyy
	    Import-Module ".\Configuration\PSExcel-master\PSExcel"
	    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'
    }
    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Initialization " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    function Hide-Console
    {
	    $consolePtr = [Console.Window]::GetConsoleWindow()
	    #0 hide
	    [Console.Window]::ShowWindow($consolePtr, 0)
    }
    function LoadXml ($global:filename)
    {
	    $XamlLoader = (New-Object System.Xml.XmlDocument)
	    $XamlLoader.Load($filename)
	    return $XamlLoader
    }
    try
    {
	    $SplashScreen = LoadXml("Configuration\XAML\SplashScreen.xaml")
	    $Splash = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $SplashScreen))
	    $SplashScreen.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{ Set-Variable -Name ($_.Name) -Value $Splash.FindName($_.Name) }
	    $SplashScreen.source = ".\Configuration\Pictures\Version.jpg"
	    $Splash.Show()
	    Start-Sleep -Seconds 2
    }
    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_SplashScreen " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    function Get-UserNameSessionIDMap ($Comp)
    {
	    $quserRes = quser /server:$comp | select -skip 1
	    if (!$quserRes) { RETURN }
	    $quCSV = @()
	    $quCSVhead = "SessionID", "UserName", "LogonTime"
	    foreach ($qur in $quserRes)
	    {
		    $qurMap = $qur.Trim().Split(" ") | ? { $_ }
		    if ($qur -notmatch " Disc   ") { $quCSV += $qurMap[2] + "|" + $qurMap[0] + "|" + $qurMap[5] + " " + $qurMap[6] }
		    else { $quCSV += $qurMap[1] + "|" + $qurMap[0] + "|" + $qurMap[4] + " " + $qurMap[5] } #disconnected sessions have no SESSIONNAME, others have ica-tcp#x
	    }
	    $quCSV | ConvertFrom-CSV -Delimiter "|" -Header $quCSVhead
    }
    function Test_DDC
    {
	    try
	    {
		    $DDC_State.source = $Grey
		    if ($DDC_TB.text -eq "")
		    {
			    $ApplicationLayer.IsEnabled = $false
			    $DDC_MB.Foreground = "Red"
			    $DDC_MB.FontSize = "20"
			    $DDC_MB.text = "Please enter a DDC."
			    $Dialog_DDC.IsOpen = $True
			    $DDC_MB_Close.add_Click({
					    $Dialog_DDC.IsOpen = $False
					    $ApplicationLayer.IsEnabled = $true
				    })
		    }
		    Elseif ([regex]::IsMatch($DDC_TB.text, $pattern))
		    {
			    $ApplicationLayer.IsEnabled = $false
			    $DDC_State.source = $Red
			    $DDC_MB.Foreground = "Red"
			    $DDC_MB.FontSize = "20"
			    $DDC_MB.text = "Please avoid special characters."
			    $Dialog_DDC.IsOpen = $True
			    $DDC_MB_Close.add_Click({
					    $Dialog_DDC.IsOpen = $False
					    $ApplicationLayer.IsEnabled = $true
				    })
		    }
		    Else
		    {
			    $ApplicationLayer.IsEnabled = $false
			    $SpinnerOverlayLayer.Visibility = "Visible"
			    $Main_Load_TB.text = "Testing DDC"
			    $Global:SyncHash_Conf = [hashtable]::Synchronized(@{
					    Form_Configuration  = $Form_Configuration
					    SpinnerOverlayLayer = $SpinnerOverlayLayer
					    ApplicationLayer    = $ApplicationLayer
					    DDC_MB			    = $DDC_MB
					    Dialog_DDC		    = $Dialog_DDC
					    DDC_State		    = $DDC_State
					    Green			    = $Green
					    Red				    = $Red
					    Grey			    = $Grey
					    DDC_TB			    = $DDC_TB.text
					    ListView_Conf	    = $ListView_Conf
				    })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_Conf", $SyncHash_Conf)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    if (!(Test-Connection $SyncHash_Conf.DDC_TB -Count 1 -ErrorAction SilentlyContinue))
						    {
							    $SyncHash_Conf.Form_Configuration.Dispatcher.Invoke([action]{
									    $SyncHash_Conf.SpinnerOverlayLayer.Visibility = "Collapsed"
									    $SyncHash_Conf.DDC_MB.Foreground = "Red"
									    $SyncHash_Conf.DDC_MB.FontSize = "20"
									    $SyncHash_Conf.DDC_MB.text = "DDC not reachable.`r`nCheck the name or be sure the DDC is accessible."
									    $SyncHash_Conf.Dialog_DDC.IsOpen = $True
									    $SyncHash_Conf.DDC_State.source = $SyncHash_Conf.Red
								    }, "Normal")
						    }
						    elseif ($SyncHash_Conf.DDC_TB -match $env:COMPUTERNAME)
						    {
							    if ((Get-Service -Name "CitrixBrokerService").Name -notmatch "CitrixBrokerService")
							    {
								    $SyncHash_Conf.ListView_Conf.Dispatcher.Invoke([Action]{ $SyncHash_Conf.DDC_TB = "" }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash_Conf.Form_Configuration.Dispatcher.Invoke([action]{
										    $SyncHash_Conf.SpinnerOverlayLayer.Visibility = "Collapsed"
										    $SyncHash_Conf.DDC_MB.Foreground = "Red"
										    $SyncHash_Conf.DDC_MB.FontSize = "20"
										    $SyncHash_Conf.DDC_MB.text = "It's not a DDC."
										    $SyncHash_Conf.Dialog_DDC.IsOpen = $True
										    $SyncHash_Conf.DDC_State.source = $SyncHash_Conf.Red
									    }, "Normal")
							    }
							    else
							    {
								    $SyncHash_Conf.Form_Configuration.Dispatcher.Invoke([action]{
										    $SyncHash_Conf.SpinnerOverlayLayer.Visibility = "Collapsed"
										    $SyncHash_Conf.ApplicationLayer.IsEnabled = $true
										    $SyncHash_Conf.DDC_State.source = $SyncHash_Conf.Green
									    }, "Normal")
								
							    }
						    }
						    else
						    {
							    if (((Invoke-Command -ComputerName $SyncHash_Conf.DDC_TB -ScriptBlock { Get-Service -Name "CitrixBrokerService" }).Name) -notmatch "CitrixBrokerService")
							    {
								    $SyncHash_Conf.ListView_Conf.Dispatcher.Invoke([Action]{ $SyncHash_Conf.DDC_TB = "" }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash_Conf.Form_Configuration.Dispatcher.Invoke([action]{
										    $SyncHash_Conf.SpinnerOverlayLayer.Visibility = "Collapsed"
										    $SyncHash_Conf.DDC_MB.Foreground = "Red"
										    $SyncHash_Conf.DDC_MB.FontSize = "20"
										    $SyncHash_Conf.DDC_MB.text = "It's not a DDC or WinRM is disabled."
										    $SyncHash_Conf.Dialog_DDC.IsOpen = $True
										    $SyncHash_Conf.DDC_State.source = $SyncHash_Conf.Red
									    }, "Normal")
							    }
							    else
							    {
								    $SyncHash_Conf.Form_Configuration.Dispatcher.Invoke([action]{
										    $SyncHash_Conf.SpinnerOverlayLayer.Visibility = "Collapsed"
										    $SyncHash_Conf.ApplicationLayer.IsEnabled = $true
										    $SyncHash_Conf.DDC_State.source = $SyncHash_Conf.Green
									    }, "Normal")
								
							    }
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Test_DDC_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $DDC_MB_Close.add_Click({
					    $Dialog_DDC.IsOpen = $False
					    $ApplicationLayer.IsEnabled = $true
				    })
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Test_DDC " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Search_User
    {
	    try
	    {
		    $datagrid_usersList.Visibility = "Visible"
		    $Search_sessions.Visibility = "Visible"
		    $datagrid_UserSessions.Visibility = "Visible"
		    $Kill_Session.Visibility = "Visible"
		    $Hide_Session.Visibility = "Visible"
		    $Shadow_Session.Visibility = "Visible"
		    $Refresh_Session.Visibility = "Visible"
		    $Grid_AllPSessions_Full.Visibility = "Collapsed"
		    $Border_AllSessions.Visibility = "Collapsed"
		    $TB_AllSessions.Visibility = "Collapsed"
		    $datagrid_usersList.ItemsSource = $null
		    $datagrid_UserSessions.ItemsSource = $null
		    $S_Sessions_Details.SelectedItem = $null
		    if ($UserName_TB.text -eq "") { Show-Dialog_Main -Foreground "Red" -Text "Please enter a username." }
		    elseif ($UserName_TB.text.Length -lt "2") { Show-Dialog_Main -Foreground "Red" -Text "Please enter at least 2 characters." }
		    else
		    {
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Searching users"
			    $Global:SyncHash_Sessions = [hashtable]::Synchronized(@{ UserName_TB = $UserName_TB.text })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_Sessions", $SyncHash_Sessions)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $Users_list = @()
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($Farm in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$Farm).DDC
							    $Users_list += Get-BrokerUser -MaxRecordCount 999999 -AdminAddress $DDC | ? { $_.FullName -match $SyncHash_Sessions.UserName_TB -or $_.UPN -match $SyncHash_Sessions.UserName_TB } | Select-Object FullName, UPN, @{ n = "Domain"; e = { $_.Name.Split('\')[0] } }, SID
						    }
						    $Users_list = $Users_list | Select-Object FullName, UPN, Domain, SID -Unique | Sort-Object FullName
						    if ($Users_List -eq $null)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No user found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($Users_List.FullName.count -eq 1)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $Users_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $Users_List_Datagrid.Add($Users_List)
							    $SyncHash.datagrid_usersList.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_usersList.ItemsSource = $Users_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
						    else
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $Users_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $Users_List_Datagrid.AddRange($Users_List)
							    $SyncHash.datagrid_usersList.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_usersList.ItemsSource = $Users_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_User_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_User " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Search_Publication
    {
	    try
	    {
		    $datagrid_publications.ItemsSource = $null
		    $datagrid_application_settings.ItemsSource = $null
		    $datagrid_application_settings_2.ItemsSource = $null
		    $datagrid_desktop_settings.ItemsSource = $null
		    $datagrid_desktop_settings_2.ItemsSource = $null
		    $listbox_desktop_tag.ItemsSource = $null
		    $S_Publications_Details.SelectedItem = $null
		    $datagrid_publications.Visibility = "Visible"
		    $Publication_settings.Visibility = "Visible"
		    $Publication_sessions.Visibility = "Visible"
		    $Publication_servers.Visibility = "Visible"
		    $Publication_access.Visibility = "Visible"
		    Publications_collapse
		    $Grid_AllPublications_Full.Visibility = "Collapsed"
		    $Border_AllPublis.Visibility = "Collapsed"
		    $TB_AllPublis.Visibility = "Collapsed"
		    if ($Publication_TB.text -eq "") { Show-Dialog_Main -Foreground "Red" -Text "Please enter a publication name." }
		    elseif ($Publication_TB.text.Length -lt "2") { Show-Dialog_Main -Foreground "Red" -Text "Please enter at least 2 characters." }
		    else
		    {
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Searching publications"
			    $Global:SyncHash_Publications = [hashtable]::Synchronized(@{ Publication_TB = $Publication_TB.text })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_Publications", $SyncHash_Publications)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $Publications_list = @()
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($Farm in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$Farm).DDC
							    $Publications_list += Get-BrokerApplication -MaxRecordCount 9999 -AdminAddress $DDC | ? { $_.ApplicationName -match $SyncHash_Publications.Publication_TB } | Select-Object  @{ n = "Name"; e = { $_.ApplicationName } }, @{ n = "Farm"; e = { $Farm } }, @{ n = "Sessions"; e = { (Get-BrokerSession -MaxRecordCount 9999 -AdminAddress $DDC -ApplicationUid $_.UID).count } }, Description, @{ n = "Type"; e = { "Application" } }, AssociatedDesktopGroupUids, @{ n = "Delivery Groups"; e = { ForEach ($AssociatedDesktopGroupUid in $_.AssociatedDesktopGroupUids) { Get-BrokerDesktopGroup -AdminAddress $DDC -Uid $AssociatedDesktopGroupUid | Select-Object -ExpandProperty name } } }, UID
							    $Publications_list += Get-BrokerEntitlementPolicyRule -MaxRecordCount 9999 -AdminAddress $DDC | ? { $_.Name -match $SyncHash_Publications.Publication_TB } | Select-Object @{ n = "Name"; e = { $_.Name } }, @{ n = "Farm"; e = { $Farm } }, @{ n = "Sessions"; e = { (Get-BrokerSession -MaxRecordCount 9999 -AdminAddress $DDC -LaunchedViaPublishedName $_.PublishedName).count } }, Description, @{ n = "Type"; e = { "Desktop" } }, DesktopGroupUid, @{ n = "Delivery Groups"; e = { Get-BrokerDesktopGroup -AdminAddress $DDC -Uid $_.DesktopGroupUid | Select-Object -ExpandProperty name } }, UID
						    }
						    $Publications_list = $Publications_list | Sort-Object Name
						    if ($Publications_list.count -eq 0)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No publication found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($Publications_list.Name.count -eq 1)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $Publications_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $Publications_List_Datagrid.Add($Publications_list)
							    $SyncHash.datagrid_publications.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_publications.ItemsSource = $Publications_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
						    else
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $Publications_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $Publications_List_Datagrid.AddRange($Publications_list)
							    $SyncHash.datagrid_publications.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_publications.ItemsSource = $Publications_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_Publication_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_Publication " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Search_VDA
    {
	    try
	    {
		    $datagrid_VDAsList.Visibility = "Visible"
		    $VDA_Details.Visibility = "Visible"
		    $VDA_Sessions.Visibility = "Visible"
		    $VDA_Publications.Visibility = "Visible"
		    $VDA_Hotfixes.Visibility = "Visible"
		    VDAs_collapse
		    $Grid_AllVDAs_Full.Visibility = "collapse"
		    $Border_AllVDAs.Visibility = "collapse"
		    $TB_AllVDAs.Visibility = "collapse"
		    $datagrid_VDAsList.ItemsSource = $null
		    $S_VDAs_Details.SelectedItem = $null
		    if ($VDA_TB.text -eq "") { Show-Dialog_Main -Foreground "Red" -Text "Please enter a VDA name." }
		    elseif ($VDA_TB.text.Length -lt "2") { Show-Dialog_Main -Foreground "Red" -Text "Please enter at least 2 characters." }
		    else
		    {
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Searching VDAs"
			    $Global:SyncHash_VDAs = [hashtable]::Synchronized(@{ VDA_TB = $VDA_TB.text })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_VDAs", $SyncHash_VDAs)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $VDAs_list = @()
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($Farm in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$Farm).DDC
							    $VDAs_list += Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $DDC | ? { $_.MachineName -match $SyncHash_VDAs.VDA_TB } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{
								    n = "Type"; e = {
									    if ($_.SessionSupport -match "MultiSession") { "Server" }
									    else { "VDI" }
								    }
							    }, @{ n = "Farm"; e = { $Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Sessions"; e = { $_.SessionCount } }, UID
						    }
						    $VDAs_list = $VDAs_list | Sort-Object "Machine Name"
						    if ($VDAs_list."Machine Name".count -eq 0)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No VDA found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($VDAs_list."Machine Name".count -eq 1)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $VDAs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $VDAs_List_Datagrid.Add($VDAs_list)
							    $SyncHash.datagrid_VDAsList.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_VDAsList.ItemsSource = $VDAs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
						    else
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $VDAs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $VDAs_List_Datagrid.AddRange($VDAs_list)
							    $SyncHash.datagrid_VDAsList.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_VDAsList.ItemsSource = $VDAs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_VDA_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_VDA " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Search_MC
    {
	    try
	    {
		    $datagrid_MCsList.Visibility = "Visible"
		    $MC_Details.Visibility = "Visible"
		    $MC_VDAs.Visibility = "Visible"
		    $MC_Sessions.Visibility = "Visible"
		    $MC_Refresh.Visibility = "Visible"
		    $Border_AllMCs.Visibility = "collapse"
		    $TB_AllMCs.Visibility = "collapse"
		    MCs_collapse
		    $datagrid_MCsList.ItemsSource = $null
		    $S_MCs_Details.SelectedItem = $null
		    if ($MC_TB.text -eq "") { Show-Dialog_Main -Foreground "Red" -Text "Please enter a Machine Catalog name." }
		    elseif ($MC_TB.text.Length -lt "2") { Show-Dialog_Main -Foreground "Red" -Text "Please enter at least 2 characters." }
		    else
		    {
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Searching Machine Catalogs"
			    $Global:SyncHash_MCs = [hashtable]::Synchronized(@{ MC_TB = $MC_TB.text })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_MCs", $SyncHash_MCs)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $MCs_list = @()
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($Farm in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$Farm).DDC
							    $MCs_list += Get-BrokerCatalog -MaxRecordCount 999999 -AdminAddress $DDC | ? { $_.Name -match $SyncHash_MCs.MC_TB } | ForEach-Object { $MC_Name = $_.Name; $_ } | Select-Object Name, @{ n = "Farm"; e = { $Farm } }, @{ n = "VDAs"; e = { $_.UsedCount + $_.AvailableCount } }, @{ n = "Sessions"; e = { (Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $DDC -CatalogName $MC_Name).count } }, UID
						    }
						    $MCs_list = $MCs_list | Sort-Object Name
						    if ($MCs_list.Name.count -eq 0)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No Machine Catalog found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($MCs_list.Name.count -eq 1)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $MCs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $MCs_List_Datagrid.Add($MCs_list)
							    $SyncHash.datagrid_MCsList.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_MCsList.ItemsSource = $MCs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
						    else
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $MCs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $MCs_List_Datagrid.AddRange($MCs_list)
							    $SyncHash.datagrid_MCsList.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_MCsList.ItemsSource = $MCs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_MC_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_MC " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Search_DG
    {
	    try
	    {
		    $datagrid_DGsList.Visibility = "Visible"
		    $DG_Details.Visibility = "Visible"
		    $DG_VDAs.Visibility = "Visible"
		    $DG_Sessions.Visibility = "Visible"
		    $DG_Publications.Visibility = "Visible"
		    $DG_Refresh.Visibility = "Visible"
		    $Border_AllDGs.Visibility = "collapse"
		    $TB_AllDGs.Visibility = "collapse"
		    DGs_collapse
		    $datagrid_DGsList.ItemsSource = $null
		    $S_DGs_Details.SelectedItem = $null
		    if ($DG_TB.text -eq "") { Show-Dialog_Main -Foreground "Red" -Text "Please enter a Delivery Group name." }
		    elseif ($DG_TB.text.Length -lt "2") { Show-Dialog_Main -Foreground "Red" -Text "Please enter at least 2 characters." }
		    else
		    {
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Global:SyncHash_DGs = [hashtable]::Synchronized(@{ DG_TB = $DG_TB.text })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_DGs", $SyncHash_DGs)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $DGs_list = @()
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($Farm in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$Farm).DDC
							    $DGs_list += Get-BrokerDesktopGroup -MaxRecordCount 999999 -AdminAddress $DDC | ? { $_.Name -match $SyncHash_DGs.DG_TB } | ForEach-Object { $DG_Name = $_.Name; $_ } | ForEach-Object { $DG_Uid = $_.Uid; $_ } | Select-Object Name, @{ n = "Farm"; e = { $Farm } }, @{ n = "VDAs"; e = { $_.TotalDesktops } }, @{ n = "Sessions"; e = { $_.Sessions } }, @{ n = "Publications"; e = { (Get-BrokerApplication -MaxRecordCount 99999 -AdminAddress $DDC | Where-Object { $_.AssociatedDesktopGroupUids -eq $DG_Uid } | Select-Object -ExpandProperty ApplicationName).count + (Get-BrokerEntitlementPolicyRule -MaxRecordCount 9999 -AdminAddress $DDC -DesktopGroupUid $DG_Uid | Select-Object -ExpandProperty Name).count } }, UID
						    }
						    $DGs_list = $DGs_list | Sort-Object Name
						    if ($DGs_list.Name.count -eq 0)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No Delivery Group found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($DGs_list.Name.count -eq 1)
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $DGs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $DGs_List_Datagrid.Add($DGs_list)
							    $SyncHash.datagrid_DGsList.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DGsList.ItemsSource = $DGs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
						    else
						    {
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
							    $DGs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $DGs_List_Datagrid.AddRange($DGs_list)
							    $SyncHash.datagrid_DGsList.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DGsList.ItemsSource = $DGs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_DG_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_DG " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function S_DGs_Details
    {
	    try
	    {
		    if ($S_DGs_Details.selectedItem -ne $null)
		    {
			    $DG_TB.text = ""
			    $Border_AllDGs.Visibility = "collapse"
			    $TB_AllDGs.Visibility = "collapse"
			    $TB_AllDGs.Text = ""
			    $datagrid_DGsList.Visibility = "Visible"
			    $Border_AllDGs.Visibility = "Visible"
			    $TB_AllDGs.Visibility = "Visible"
			    $DG_Details.Visibility = "Visible"
			    $DG_VDAs.Visibility = "Visible"
			    $DG_Sessions.Visibility = "Visible"
			    $DG_Publications.Visibility = "Visible"
			    $DG_Refresh.Visibility = "Visible"
			    DGs_collapse
			    $datagrid_DGsList.ItemsSource = $null
			    $Farm = $S_DGs_Details.selectedItem
			    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Global:SyncHash_AllDGs_list = [hashtable]::Synchronized(@{
					    Farm	  = $Farm
					    DDC	      = $DDC
					    Farm_List = $SyncHash.Farm_List
				    })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_AllDGs_list", $SyncHash_AllDGs_list)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $DGs_list = @()
						    $Total_DGs = @()
						    if ($SyncHash_AllDGs_list.Farm -eq "All Farms")
						    {
							    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
							    foreach ($item in $SyncHash.Farm_List)
							    {
								    $DDC = ($SyncHash.$item).DDC
								    $DGs_list += Get-BrokerDesktopGroup -MaxRecordCount 999999 -AdminAddress $DDC | ForEach-Object { $DG_Name = $_.Name; $_ } | ForEach-Object { $DG_Uid = $_.Uid; $_ } | Select-Object Name, @{ n = "Farm"; e = { $item } }, @{ n = "VDAs"; e = { $_.TotalDesktops } }, @{ n = "Sessions"; e = { $_.Sessions } },
																																																					    @{ n = "Publications"; e = { (Get-BrokerApplication -MaxRecordCount 99999 -AdminAddress $DDC | Where-Object { $_.AssociatedDesktopGroupUids -eq $DG_Uid } | Select-Object -ExpandProperty ApplicationName).count + (Get-BrokerEntitlementPolicyRule -MaxRecordCount 9999 -AdminAddress $DDC -DesktopGroupUid $DG_Uid | Select-Object -ExpandProperty Name).count } }, UID
							    }
							    $DGs_list = $DGs_list | Sort-Object Name
							    $Total_DGs = $DGs_list.Name.count
						    }
						    else
						    {
							    $DGs_list = Get-BrokerDesktopGroup -MaxRecordCount 999999 -AdminAddress $SyncHash_AllDGs_list.DDC | ForEach-Object { $DG_Name = $_.Name; $_ } | ForEach-Object { $DG_Uid = $_.Uid; $_ } | Select-Object Name, @{ n = "Farm"; e = { $SyncHash_AllDGs_list.Farm } },
																																																									    @{ n = "VDAs"; e = { $_.TotalDesktops } }, @{ n = "Sessions"; e = { $_.Sessions } },
																																																									    @{ n = "Publications"; e = { (Get-BrokerApplication -MaxRecordCount 99999 -AdminAddress $SyncHash_AllDGs_list.DDC | Where-Object { $_.AssociatedDesktopGroupUids -eq $DG_Uid } | Select-Object -ExpandProperty ApplicationName).count + (Get-BrokerEntitlementPolicyRule -MaxRecordCount 9999 -AdminAddress $SyncHash_AllDGs_list.DDC -DesktopGroupUid $DG_Uid | Select-Object -ExpandProperty Name).count } }, UID
							    $DGs_list = $DGs_list | Sort-Object Name
							    $Total_DGs = $DGs_list.Name.count
						    }
						    if ($DGs_list.name.count -eq 0)
						    {
							    $SyncHash.TB_AllDGs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllDGs.Text = "Total Delivery Groups : $Total_DGs" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Border_AllDGs.Visibility = "Visible"
									    $SyncHash.TB_AllDGs.Visibility = "Visible"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No VDA found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($DGs_list.name.count -eq 1)
						    {
							    $AllDGs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $AllDGs_List_Datagrid.Add($DGs_list)
							    $SyncHash.TB_AllDGs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllDGs.Text = "Total Delivery Groups : $Total_DGs" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.datagrid_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DGsList.ItemsSource = $AllDGs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Border_AllDGs.Visibility = "Visible"
									    $SyncHash.TB_AllDGs.Visibility = "Visible"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
						    else
						    {
							    $AllDGs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $AllDGs_List_Datagrid.AddRange($DGs_list)
							    $SyncHash.TB_AllDGs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllDGs.Text = "Total Delivery Groups : $Total_DGs" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.datagrid_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DGsList.ItemsSource = $AllDGs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Border_AllDGs.Visibility = "Visible"
									    $SyncHash.TB_AllDGs.Visibility = "Visible"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_DGs_Details_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_DGs_Details " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function S_MCs_Details
    {
	    try
	    {
		    if ($S_MCs_Details.selectedItem -ne $null)
		    {
			    $MC_TB.text = ''
			    $Border_AllMCs.Visibility = "collapse"
			    $TB_AllMCs.Visibility = "collapse"
			    $TB_AllMCs.Text = ""
			    $datagrid_MCsList.Visibility = "Visible"
			    $Border_AllMCs.Visibility = "Visible"
			    $TB_AllMCs.Visibility = "Visible"
			    $MC_Details.Visibility = "Visible"
			    $MC_VDAs.Visibility = "Visible"
			    $MC_Sessions.Visibility = "Visible"
			    $MC_Refresh.Visibility = "Visible"
			    MCs_collapse
			    $datagrid_MCsList.ItemsSource = $null
			    $Farm = $S_MCs_Details.selectedItem
			    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Global:SyncHash_AllMCs_list = [hashtable]::Synchronized(@{
					    Farm	  = $Farm
					    DDC	      = $DDC
					    Farm_List = $SyncHash.Farm_List
				    })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_AllMCs_list", $SyncHash_AllMCs_list)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $MCs_list = @()
						    $Total_MCs = @()
						    if ($SyncHash_AllMCs_list.Farm -eq "All Farms")
						    {
							    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
							    foreach ($item in $SyncHash.Farm_List)
							    {
								    $DDC = ($SyncHash.$item).DDC
								    $MCs_list += Get-BrokerCatalog -MaxRecordCount 999999 -AdminAddress $DDC | ForEach-Object { $MC_Name = $_.Name; $_ } | Select-Object Name, @{ n = "Farm"; e = { $item } }, @{ n = "VDAs"; e = { $_.UsedCount + $_.AvailableCount } }, @{ n = "Sessions"; e = { (Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $DDC -CatalogName $MC_Name).count } }, UID
							    }
							    $MCs_list = $MCs_list | Sort-Object Name
							    $Total_MCs = $MCs_list.Name.count
						    }
						    else
						    {
							    $MCs_list = Get-BrokerCatalog -MaxRecordCount 999999 -AdminAddress $SyncHash_AllMCs_list.DDC | ForEach-Object { $MC_Name = $_.Name; $_ } | Select-Object Name, @{ n = "Farm"; e = { $SyncHash_AllMCs_list.Farm } }, @{ n = "VDAs"; e = { $_.UsedCount + $_.AvailableCount } }, @{ n = "Sessions"; e = { (Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $SyncHash_AllMCs_list.DDC -CatalogName $MC_Name).count } }, UID
							    $MCs_list = $MCs_list | Sort-Object Name
							    $Total_MCs = $MCs_list.Name.count
						    }
						    if ($MCs_list.name.count -eq 0)
						    {
							    $SyncHash.TB_AllMCs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllMCs.Text = "Total Machine Catalogs : $Total_MCs" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Border_AllMCs.Visibility = "Visible"
									    $SyncHash.TB_AllMCs.Visibility = "Visible"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No VDA found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($MCs_list.name.count -eq 1)
						    {
							    $AllMCs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $AllMCs_List_Datagrid.Add($MCs_list)
							    $SyncHash.TB_AllMCs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllMCs.Text = "Total Machine Catalogs : $Total_MCs" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.datagrid_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_MCsList.ItemsSource = $AllMCs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Border_AllMCs.Visibility = "Visible"
									    $SyncHash.TB_AllMCs.Visibility = "Visible"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
						    else
						    {
							    $AllMCs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $AllMCs_List_Datagrid.AddRange($MCs_list)
							    $SyncHash.TB_AllMCs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllMCs.Text = "Total Machine Catalogs : $Total_MCs" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.datagrid_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_MCsList.ItemsSource = $AllMCs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Border_AllMCs.Visibility = "Visible"
									    $SyncHash.TB_AllMCs.Visibility = "Visible"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_MCs_Details_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_MCs_Details " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function LoadXml ($global:filename)
    {
	    $XamlLoader = (New-Object System.Xml.XmlDocument)
	    $XamlLoader.Load($filename)
	    return $XamlLoader
    }
    function Refresh_Session
    {
	    try
	    {
		    $datagrid_UserSessions.ItemsSource = $null
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_Sessions_list = [hashtable]::Synchronized(@{
				    user   = $datagrid_usersList.SelectedItem.fullname
				    domain = $datagrid_usersList.SelectedItem.domain
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_Sessions_list", $SyncHash_Sessions_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $Users_sessions = @()
					    $user = $SyncHash_Sessions_list.user
					    $domain = $SyncHash_Sessions_list.domain
					    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
					    foreach ($Farm in $SyncHash.Farm_List)
					    {
						    $DDC = ($SyncHash.$Farm).DDC
						    $Users_sessions += Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $DDC -UserName $domain\$user | Select-Object @{ n = "User"; e = { $_.UserFullName } }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { $Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
							    n = "Type"; e = {
								    if ($_.SessionSupport -match "MultiSession") { "Server" }
								    else { "VDI" }
							    }
						    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID
					    }
					    if ($Users_sessions.count -eq 0)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No session found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $SyncHash.datagrid_UserSessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_UserSessions.ItemsSource = $Users_sessions }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Session_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Session " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Refresh_AllSessions
    {
	    try
	    {
		    $datagrid_AllSessions.ItemsSource = $null
		    $TextBox_AllSessions.Text = ""
		    $TB_AllSessions.Text = ""
		    $Total_Sessions = $null
		    $Active_Sessions = $null
		    $Disconnected_Sessions = $null
		    $Connected_Sessions = $null
		    $Farm = $S_Sessions_Details.selectedItem
		    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_AllSessions_list = [hashtable]::Synchronized(@{
				    Farm	  = $Farm
				    DDC	      = $DDC
				    Farm_List = $SyncHash.Farm_List
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_AllSessions_list", $SyncHash_AllSessions_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $AllSessions_list = @()
					    if ($SyncHash_AllSessions_list.Farm -eq "All Farms")
					    {
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($item in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$item).DDC
							    $AllSessions_list += Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $DDC | Select-Object @{
								    n = "User"; e = {
									    if ($_.UserFullName -eq $null) { ".no data" }
									    else { $_.UserFullName }
								    }
							    }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { ($SyncHash.$item).Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
								    n = "Type"; e = {
									    if ($_.SessionSupport -match "MultiSession") { "Server" }
									    else { "VDI" }
								    }
							    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID
						    }
						    $AllSessions_list = $AllSessions_list | Sort-Object user
						    $Total_Sessions = $AllSessions_list.User.count
						    $Active_Sessions = ($AllSessions_list | Where-Object { $_.SessionState -eq "Active" }).User.count
						    $Disconnected_Sessions = ($AllSessions_list | Where-Object { $_.SessionState -eq "Disconnected" }).User.count
						    $Connected_Sessions = ($AllSessions_list | Where-Object { $_.SessionState -eq "Connected" }).User.count
						    $Simple_AllSessions_List = $AllSessions_list.User
						    if ($Total_Sessions -eq 0) { $Simple_AllSessions_List_String = $null }
						    else { $Simple_AllSessions_List_String = [string]::Join([Environment]::NewLine, $Simple_AllSessions_List) }
					    }
					    else
					    {
						    $AllSessions_list = Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $SyncHash_AllSessions_list.DDC | Select-Object @{
							    n = "User"; e = {
								    if ($_.UserFullName -eq $null) { ".no data" }
								    else { $_.UserFullName }
							    }
						    }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { $SyncHash_AllSessions_list.Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
							    n = "Type"; e = {
								    if ($_.SessionSupport -match "MultiSession") { "Server" }
								    else { "VDI" }
							    }
						    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID
						    $AllSessions_list = $AllSessions_list | Sort-Object user
						    $Total_Sessions = $AllSessions_list.User.count
						    $Active_Sessions = ($AllSessions_list | Where-Object { $_.SessionState -eq "Active" }).User.count
						    $Disconnected_Sessions = ($AllSessions_list | Where-Object { $_.SessionState -eq "Disconnected" }).User.count
						    $Connected_Sessions = ($AllSessions_list | Where-Object { $_.SessionState -eq "Connected" }).User.count
						    $Simple_AllSessions_List = $AllSessions_list.User
						    if ($Total_Sessions -eq 0) { $Simple_AllSessions_List_String = $null }
						    else { $Simple_AllSessions_List_String = [string]::Join([Environment]::NewLine, $Simple_AllSessions_List) }
					    }
					    if ($Connected_Sessions -eq 0) { $Connected_Sessions_Color = "Green" }
					    else { $Connected_Sessions_Color = "Red" }
					    if ($Total_Sessions -eq 0)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Border_AllSessions.Visibility = "Visible"
								    $SyncHash.TB_AllSessions.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No session found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
						    $SyncHash.TB_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllSessions.Text = "Total sessions : $Total_Sessions    Active Sessions : $Active_Sessions`r`nDisconnected Sessions : $Disconnected_Sessions    Connected Sessions : " }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllSessions.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = $Connected_Sessions; Foreground = $Connected_Sessions_Color })) }, [Windows.Threading.DispatcherPriority]::Normal)
						
					    }
					    elseif ($Total_Sessions -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_AllPSessions_Full.Visibility = "Visible"
								    $SyncHash.Border_AllSessions.Visibility = "Visible"
								    $SyncHash.TB_AllSessions.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $Allsessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $Allsessions_List_Datagrid.Add($Allsessions_list)
						    $SyncHash.datagrid_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_AllSessions.ItemsSource = $Allsessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllSessions.Text = "Total sessions : $Total_Sessions    Active Sessions : $Active_Sessions`r`nDisconnected Sessions : $Disconnected_Sessions    Connected Sessions : " }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllSessions.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = $Connected_Sessions; Foreground = $Connected_Sessions_Color })) }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_AllSessions.text = $Simple_AllSessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_AllPSessions_Full.Visibility = "Visible"
								    $SyncHash.Border_AllSessions.Visibility = "Visible"
								    $SyncHash.TB_AllSessions.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $Allsessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $Allsessions_List_Datagrid.AddRange($Allsessions_list)
						    $SyncHash.datagrid_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_AllSessions.ItemsSource = $Allsessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllSessions.Text = "Total sessions : $Total_Sessions    Active Sessions : $Active_Sessions`r`nDisconnected Sessions : $Disconnected_Sessions    Connected Sessions : " }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllSessions.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = $Connected_Sessions; Foreground = $Connected_Sessions_Color })) }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_AllSessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_AllSessions.text = $Simple_AllSessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllSessions_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllSessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Refresh_PubliSession
    {
	    try
	    {
		    $datagrid_PubliSessions.ItemsSource = $null
		    $Farm = $datagrid_publications.selecteditem.Farm
		    $DDC = ($SyncHash.$Farm).DDC
		    $UID = $datagrid_publications.selecteditem.UID
		    $Type = $datagrid_publications.selecteditem.Type
		    $Name = $datagrid_publications.selecteditem.Name
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_PubliSessions_list = [hashtable]::Synchronized(@{
				    DDC  = $DDC
				    UID  = $UID
				    Type = $Type
				    Name = $Name
				    Farm = $Farm
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_PubliSessions_list", $SyncHash_PubliSessions_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    if ($SyncHash_PubliSessions_list.Type -eq "Application")
					    {
						    $App = Get-BrokerApplication -AdminAddress $SyncHash_PubliSessions_list.DDC -Uid $SyncHash_PubliSessions_list.UID
						    $Sessions = Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $SyncHash_PubliSessions_list.DDC -ApplicationUid $App.uid | Select-Object @{
							    n = "User"; e = {
								    if ($_.UserFullName -eq $null) { ".no data" }
								    else { $_.UserFullName }
							    }
						    }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { $SyncHash_PubliSessions_list.Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
							    n = "Type"; e = {
								    if ($_.SessionSupport -match "MultiSession") { "Server" }
								    else { "VDI" }
							    }
						    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID | Sort-Object user
						    $Sessions = $Sessions | Sort-Object User
						    $Total_Publi_Sessions = $Sessions.User.count
						    $Simple_Publi_Sessions_List = $Sessions.User
						    if ($Total_Publi_Sessions -eq 0) { $Simple_Publi_Sessions_List_String = $null }
						    else { $Simple_Publi_Sessions_List_String = [string]::Join([Environment]::NewLine, $Simple_Publi_Sessions_List) }
					    }
					    else
					    {
						    $App = Get-BrokerEntitlementPolicyRule -AdminAddress $SyncHash_PubliSessions_list.DDC -Uid $SyncHash_PubliSessions_list.UID -ErrorAction SilentlyContinue
						    $Sessions = Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $SyncHash_PubliSessions_list.DDC -LaunchedViaPublishedName $App.PublishedName | Select-Object @{
							    n = "User"; e = {
								    if ($_.UserFullName -eq $null) { ".no data" }
								    else { $_.UserFullName }
							    }
						    }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { $SyncHash_PubliSessions_list.Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
							    n = "Type"; e = {
								    if ($_.SessionSupport -match "MultiSession") { "Server" }
								    else { "VDI" }
							    }
						    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID | Sort-Object user
						    $Sessions = $Sessions | Sort-Object User
						    $Total_Publi_Sessions = $Sessions.User.count
						    $Simple_Publi_Sessions_List = $Sessions.User
						    if ($Total_Publi_Sessions -eq 0) { $Simple_Publi_Sessions_List_String = $null }
						    else { $Simple_Publi_Sessions_List_String = [string]::Join([Environment]::NewLine, $Simple_Publi_Sessions_List) }
					    }
					    if ($Total_Publi_Sessions -eq 0)
					    {
						    $SyncHash.TextBox_TotalPubliSessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalPubliSessions.text = "Total = $Total_Publi_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No session found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_Publi_Sessions -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_PubliSessions_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $Sessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $Sessions_List_Datagrid.Add($Sessions)
						    $SyncHash.datagrid_PubliSessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_PubliSessions.ItemsSource = $Sessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalPubliSessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalPubliSessions.text = "Total = $Total_Publi_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_PubliSessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_PubliSessions.text = $Simple_Publi_Sessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_PubliSessions_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $Sessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $Sessions_List_Datagrid.AddRange($Sessions)
						    $SyncHash.datagrid_PubliSessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_PubliSessions.ItemsSource = $Sessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalPubliSessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalPubliSessions.text = "Total = $Total_Publi_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_PubliSessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_PubliSessions.text = $Simple_Publi_Sessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_PubliSession_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_PubliSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Refresh_AllPublis
    {
	    try
	    {
		    $datagrid_AllPublications.ItemsSource = $null
		    $TextBox_AllPublications.Text = ""
		    $TB_AllPublis.Text = ""
		    $Total_Publis = $null
		    $Total_Applis = $null
		    $Total_Desktop = $null
		    $Disabled_Publis = $null
		    $Hidden_Publis = $null
		    $Farm = $S_Publications_Details.selectedItem
		    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_AllPublis_list = [hashtable]::Synchronized(@{
				    Farm	  = $Farm
				    DDC	      = $DDC
				    Farm_List = $SyncHash.Farm_List
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_AllPublis_list", $SyncHash_AllPublis_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $AllPublis_list = @()
					    $Total_Publis = @()
                        $Applis = @()
                        $Desktops = @()
					    if ($SyncHash_AllPublis_list.Farm -eq "All Farms")
					    {
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($item in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$item).DDC
							    $Applis += Get-BrokerApplication -MaxRecordCount 999999 -AdminAddress $DDC | Select-Object @{ n = "Application Name"; e = { $_.ApplicationName } }, @{ n = "Published Name"; e = { $_.PublishedName } }, @{ n = "Browser Name"; e = { $_.BrowserName } }, @{ n = "Farm"; e = { $item } },
																													       Enabled, Visible, @{ n = "Type"; e = { "Application" } }, @{ n = "Command Line Executable"; e = { $_.CommandLineExecutable } }, @{ n = "Command Line Arguments"; e = { $_.CommandLineArguments } }, @{ n = "Working Directory"; e = { $_.WorkingDirectory } }, @{ n = "Description"; e = { $_.Description } }, @{ n = "Client Folder"; e = { $_.ClientFolder } },
																													       @{ n = "Console Folder"; e = { $_.AdminFolderName } }, @{ n = "Delivery Groups"; e = { ForEach ($AssociatedDesktopGroupUid in $_.AssociatedDesktopGroupUids) { Get-BrokerDesktopGroup -AdminAddress $DDC -Uid $AssociatedDesktopGroupUid | Select-Object -ExpandProperty name } } },
																													       @{ n = "Application Groups"; e = { ForEach ($AssociatedApplicationGroupUid in $_.AssociatedApplicationGroupUids) { Get-BrokerApplicationGroup -AdminAddress $DDC -Uid $AssociatedApplicationGroupUid | Select-Object -ExpandProperty name } } }, @{ n = "Access"; e = { $_.AssociatedUserNames -join "," } }, UID
							    $Desktops += Get-BrokerEntitlementPolicyRule -MaxRecordCount 999999 -AdminAddress $DDC | Select-Object @{ n = "Application Name"; e = { $_.Name } }, @{ n = "Published Name"; e = { $_.PublishedName } }, @{ n = "Browser Name"; e = { $_.BrowserName } }, @{ n = "Farm"; e = { $item } }, Enabled, @{ n = "Visible"; e = { "N/A" } },
																																       @{ n = "Type"; e = { "Desktop" } }, @{ n = "Command Line Executable"; e = { "N/A" } }, @{ n = "Command Line Arguments"; e = { "N/A" } }, @{ n = "Working Directory"; e = { "N/A" } }, @{ n = "Description"; e = { $_.Description } }, @{ n = "Client Folder"; e = { "N/A" } }, @{ n = "Console Folder"; e = { "N/A" } },
																																       @{ n = "Delivery Groups"; e = { Get-BrokerDesktopGroup -AdminAddress $DDC -Uid $_.DesktopGroupUid | Select-Object -ExpandProperty name } }, @{ n = "Application Groups"; e = { "N/A" } }, @{ n = "Access"; e = { $_.IncludedUsers.name -join ', ' } }, UID
						    }
						    $AllPublis_list = $Applis + $Desktops
						    $AllPublis_list = $AllPublis_list | Sort-Object "Application Name"
						    $Total_Applis = $Applis."Application Name".count
						    $Total_Desktops = $Desktops."Application Name".count
						    $Total_Publis = $Total_Applis + $Total_Desktops
						    $Disabled_Publis = ($Applis | Where-Object { $_.Enabled -eq $False })."Application Name".count
						    $Hidden_Publis = ($Applis | Where-Object { $_.Visible -eq $False })."Application Name".count
						    $Disabled_Desktops = ($Desktops | Where-Object { $_.Enabled -eq $False })."Application Name".count
						    $Simple_Applis_List = $AllPublis_list."Application Name"
						    if ($Total_Publis -eq 0) { $Simple_Applis_List_String = $null }
						    else { $Simple_Applis_List_String = [string]::Join([Environment]::NewLine, $Simple_Applis_List) }
					    }
					    else
					    {
						    $Applis += Get-BrokerApplication -MaxRecordCount 999999 -AdminAddress $SyncHash_AllPublis_list.DDC | Select-Object @{ n = "Application Name"; e = { $_.ApplicationName } }, @{ n = "Published Name"; e = { $_.PublishedName } }, @{ n = "Browser Name"; e = { $_.BrowserName } }, @{ n = "Farm"; e = { $SyncHash_AllPublis_list.Farm } },
																																		       Enabled, Visible, @{ n = "Type"; e = { "Application" } }, @{ n = "Command Line Executable"; e = { $_.CommandLineExecutable } }, @{ n = "Command Line Arguments"; e = { $_.CommandLineArguments } }, @{ n = "Working Directory"; e = { $_.WorkingDirectory } }, @{ n = "Description"; e = { $_.Description } }, @{ n = "Client Folder"; e = { $_.ClientFolder } },
																																		       @{ n = "Console Folder"; e = { $_.AdminFolderName } }, @{ n = "Delivery Groups"; e = { ForEach ($AssociatedDesktopGroupUid in $_.AssociatedDesktopGroupUids) { Get-BrokerDesktopGroup -AdminAddress $SyncHash_AllPublis_list.DDC -Uid $AssociatedDesktopGroupUid | Select-Object -ExpandProperty name } } },
																																		       @{ n = "Application Groups"; e = { ForEach ($AssociatedApplicationGroupUid in $_.AssociatedApplicationGroupUids) { Get-BrokerApplicationGroup -AdminAddress $SyncHash_AllPublis_list.DDC -Uid $AssociatedApplicationGroupUid | Select-Object -ExpandProperty name } } }, @{ n = "Access"; e = { $_.AssociatedUserNames -join "," } }, UID
						    $Desktops += Get-BrokerEntitlementPolicyRule -MaxRecordCount 999999 -AdminAddress $SyncHash_AllPublis_list.DDC | Select-Object @{ n = "Application Name"; e = { $_.Name } }, @{ n = "Published Name"; e = { $_.PublishedName } }, @{ n = "Browser Name"; e = { $_.BrowserName } }, @{ n = "Farm"; e = { $SyncHash_AllPublis_list.Farm } }, Enabled, @{ n = "Visible"; e = { "N/A" } },
																																					       @{ n = "Type"; e = { "Desktop" } }, @{ n = "Command Line Executable"; e = { "N/A" } }, @{ n = "Command Line Arguments"; e = { "N/A" } }, @{ n = "Working Directory"; e = { "N/A" } }, @{ n = "Description"; e = { $_.Description } }, @{ n = "Client Folder"; e = { "N/A" } }, @{ n = "Console Folder"; e = { "N/A" } },
																																					       @{ n = "Delivery Groups"; e = { Get-BrokerDesktopGroup -AdminAddress $SyncHash_AllPublis_list.DDC -Uid $_.DesktopGroupUid | Select-Object -ExpandProperty name } }, @{ n = "Application Groups"; e = { "N/A" } }, @{ n = "Access"; e = { $_.IncludedUsers.name -join ', ' } }, UID
						    $AllPublis_list = $Applis + $Desktops
						    $AllPublis_list = $AllPublis_list | Sort-Object "Application Name"
						    $Total_Applis = $Applis."Application Name".count
						    $Total_Desktops = $Desktops."Application Name".count
						    $Total_Publis = $Total_Applis + $Total_Desktops
						    $Disabled_Publis = ($Applis | Where-Object { $_.Enabled -eq $False })."Application Name".count
						    $Hidden_Publis = ($Applis | Where-Object { $_.Visible -eq $False })."Application Name".count
						    $Disabled_Desktops = ($Desktops | Where-Object { $_.Enabled -eq $False })."Application Name".count
						    $Simple_Applis_List = $AllPublis_list."Application Name"
						    if ($Total_Publis -eq 0) { $Simple_Applis_List_String = $null }
						    else { $Simple_Applis_List_String = [string]::Join([Environment]::NewLine, $Simple_Applis_List) }
					    }
					    if ($Total_Publis -eq 0)
					    {
						    $SyncHash.TB_AllPublis.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllPublis.Text = "Total Publications : $Total_Publis    Total Applications : $Total_Applis    Total Desktops : $Total_Desktops`r`nPublications Disabled : $Disabled_Publis    Publications Hidden : $Hidden_Publis    Desktops Disabled : $Disabled_Desktops" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No publication found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_Publis -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_AllPublications_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $AllPublis_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $AllPublis_List_Datagrid.Add($AllPublis_list)
						    $SyncHash.datagrid_AllPublications.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_AllPublications.ItemsSource = $AllPublis_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllPublis.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllPublis.Text = "Total Publications : $Total_Publis    Total Applications : $Total_Applis    Total Desktops : $Total_Desktops`r`nPublications Disabled : $Disabled_Publis    Publications Hidden : $Hidden_Publis    Desktops Disabled : $Disabled_Desktops" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_AllPublications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_AllPublications.text = $Simple_Applis_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_AllPublications_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $AllPublis_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $AllPublis_List_Datagrid.AddRange($AllPublis_list)
						    $SyncHash.datagrid_AllPublications.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_AllPublications.ItemsSource = $AllPublis_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllPublis.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllPublis.Text = "Total Publications : $Total_Publis    Total Applications : $Total_Applis    Total Desktops : $Total_Desktops`r`nPublications Disabled : $Disabled_Publis    Publications Hidden : $Hidden_Publis    Desktops Disabled : $Disabled_Desktops" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_AllPublications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_AllPublications.text = $Simple_Applis_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllPublis_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Refresh_VDA
    {
	    try
	    {
		    $datagrid_VDA_settings.ItemsSource = $null
		    $Farm = $datagrid_VDAsList.selecteditem.farm
		    $UID = $datagrid_VDAsList.selecteditem.uid
		    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
		    $DDC = ($SyncHash.$Farm).DDC
		    $VDA = Get-BrokerMachine -AdminAddress $DDC -Uid $UID
		    $VDA_Registration = $VDA.RegistrationState
		    If ($VDA.InMaintenanceMode -eq "True") { $VDA_Maintenance = "Enable" }
		    else { $VDA_Maintenance = "Disable" }
		    $VDA_Power = $VDA.PowerState
		    $VDA_Farm = $Farm
		    $VDA_OS = $VDA.OSType
		    $VDA_IP = $VDA.IPAddress
		    $VDA_DG = $VDA.DesktopGroupName
		    $VDA_MC = $VDA.CatalogName
		    $VDA_Load = $VDA.LoadIndex
		    $VDA_Agent = $VDA.AgentVersion
		    $VDA_Provisioning = $VDA.ProvisioningType
		    $Applications_Groups = @()
		    if ($VDA.PowerState -eq "On")
		    {
			    if (Test-Connection -Count 1 -quiet $Name) { $VDA_BootTime = Get-CimInstance -ComputerName ([System.Net.Dns]::GetHostByName($Name)).HostName -ClassName Win32_OperatingSystem | Select -ExpandProperty LastBootUpTime }
			    else { $VDA_BootTime = "N/A" }
		    }
		    else { $VDA_BootTime = "N/A" }
		    If (($VDA.Tags).count -ne 0)
		    {
			    $VDA_TAG = $VDA.Tags | Out-String
			    foreach ($Tag in $VDA.Tags) { $Applications_Groups += Get-BrokerApplicationGroup -AdminAddress VWC2APP141 -RestrictToTag $Tag | Select-Object Name }
		    }
		    else { $VDA_TAG = "No Tag" }
		    if ($Applications_Groups.Name.count -eq "0") { $VDA_AG = "0" }
		    else { $VDA_AG = $Applications_Groups.Name | Out-String }
		    If (($VDA.Tags).count -ne 0)
		    {
			    $Applications_Groups = @()
			    foreach ($Tag in $VDA.Tags) { $Applications_Groups += Get-BrokerApplicationGroup -AdminAddress VWC2APP141 -RestrictToTag $Tag | Select-Object Name }
		    }
		    $VDA_AG = $Applications_Groups.Name | Out-String
		    $datagrid_VDA_settings.ItemsSource = @(
			    [PSCustomObject]@{ Column1Header = "Registration State"; Column2Data = $VDA_Registration; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Maintenance State"; Column2Data = $VDA_Maintenance; IsReadOnly = $true }
			    [PSCustomObject]@{ Column1Header = "Power State"; Column2Data = $VDA_Power; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Farm"; Column2Data = $VDA_Farm; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "OS Type"; Column2Data = $VDA_OS; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "IP Address"; Column2Data = $VDA_IP; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Delevery Group"; Column2Data = $VDA_DG; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Machine Catalog"; Column2Data = $VDA_MC; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Load"; Column2Data = $VDA_Load; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Agent Version"; Column2Data = $VDA_Agent; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Provisioning Type"; Column2Data = $VDA_Provisioning; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Last Boot Time"; Column2Data = $VDA_BootTime; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Tags"; Column2Data = $VDA_TAG; IsReadOnly = $true },
			    [PSCustomObject]@{ Column1Header = "Applications Groups"; Column2Data = $VDA_AG; IsReadOnly = $true }
		    )
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_VDA " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Refresh_VDASession
    {
	    try
	    {
		    $datagrid_VDA_sessions.ItemsSource = $null
		    $Farm = $datagrid_VDAsList.selecteditem.farm
		    $UID = $datagrid_VDAsList.selecteditem.uid
		    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
		    $DDC = ($SyncHash.$Farm).DDC
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Load_TB.Text = "Refreshing sessions"
		    $Global:SyncHash_VDASessions_list = [hashtable]::Synchronized(@{
				    DDC  = $DDC
				    UID  = $UID
				    Name = $Name
				    Farm = $Farm
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_VDASessions_list", $SyncHash_VDASessions_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $VDA_Sessions_List = Get-BrokerSession -MaxRecordCount 99999 -AdminAddress $SyncHash_VDASessions_list.DDC -MachineUid $SyncHash_VDASessions_list.UID | Select-Object @{
						    n = "User"; e = {
							    if ($_.UserFullName -eq $null) { ".no data" }
							    else { $_.UserFullName }
						    }
					    }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { $SyncHash_VDASessions_list.Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
						    n = "Type"; e = {
							    if ($_.SessionSupport -match "MultiSession") { "Server" }
							    else { "VDI" }
						    }
					    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID
					    if ($VDA_Sessions_List.count -eq 0)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No session found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($VDA_Sessions_List.User.count -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $VDASessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $VDASessions_List_Datagrid.Add($VDA_Sessions_List)
						    $SyncHash.datagrid_VDA_sessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_VDA_sessions.ItemsSource = $VDASessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $VDASessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $VDASessions_List_Datagrid.AddRange($VDA_Sessions_List)
						    $SyncHash.datagrid_VDA_sessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_VDA_sessions.ItemsSource = $VDASessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_VDASession_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_VDASession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Refresh_AllVDAs
    {
	    try
	    {
		    $Farm = $S_VDAs_Details.selectedItem
		    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
		    $datagrid_AllVDAs.ItemsSource = $null
		    $TextBox_AllVDAs.Text = ""
		    $TB_AllVDAs.Text = ""
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_AllVDAs_list = [hashtable]::Synchronized(@{
				    Farm	  = $Farm
				    DDC	      = $DDC
				    Farm_List = $SyncHash.Farm_List
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_AllVDAs_list", $SyncHash_AllVDAs_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $AllVDAs_list = @()
					    $Total_VDAs = @()
					    if ($SyncHash_AllVDAs_list.Farm -eq "All Farms")
					    {
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($item in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$item).DDC
							    $VDAs_list += Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $DDC | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $item } }, @{
								    n = "Type"; e = {
									    if ($_.SessionSupport -match "MultiSession") { "Server" }
									    else { "VDI" }
								    }
							    },
																													      @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } },
																													      @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
						    }
						    $VDAs_list = $VDAs_list | Sort-Object "Machine Name"
						    $Total_VDAs = $VDAs_list."Machine Name".count
						    $Total_Servers = ($VDAs_list | Where-Object { $_.Type -eq "Server" })."Machine Name".count
						    $Total_VDIs = ($VDAs_list | Where-Object { $_.Type -eq "VDI" })."Machine Name".count
						    $Total_PoweredOff = ($VDAs_list | Where-Object { $_."Power State" -eq "Off" })."Machine Name".count
						    $Total_Maintenance = ($VDAs_list | Where-Object { $_."Maintenance State" -eq $true })."Machine Name".count
						    $Total_Unregistered = ($VDAs_list | Where-Object { $_."Registration State" -eq "Unregistered" })."Machine Name".count
						    $Simple_VDAs_List = $VDAs_list."Machine Name"
						    $Simple_VDAs_List_String = [string]::Join([Environment]::NewLine, $Simple_VDAs_List)
					    }
					    else
					    {
						    $VDAs_list = Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $SyncHash_AllVDAs_list.DDC | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $SyncHash_AllVDAs_list.Farm } }, @{
							    n = "Type"; e = {
								    if ($_.SessionSupport -match "MultiSession") { "Server" }
								    else { "VDI" }
							    }
						    },
																																	       @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } },
																																	       @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
						    $VDAs_list = $VDAs_list | Sort-Object "Machine Name"
						    $Total_VDAs = $VDAs_list."Machine Name".count
						    $Total_Servers = ($VDAs_list | Where-Object { $_.Type -eq "Server" })."Machine Name".count
						    $Total_VDIs = ($VDAs_list | Where-Object { $_.Type -eq "VDI" })."Machine Name".count
						    $Total_PoweredOff = ($VDAs_list | Where-Object { $_."Power State" -eq "Off" })."Machine Name".count
						    $Total_Maintenance = ($VDAs_list | Where-Object { $_."Maintenance State" -eq $true })."Machine Name".count
						    $Total_Unregistered = ($VDAs_list | Where-Object { $_."Registration State" -eq "Unregistered" })."Machine Name".count
						    $Simple_VDAs_List = $VDAs_list."Machine Name"
						    $Simple_VDAs_List_String = [string]::Join([Environment]::NewLine, $Simple_VDAs_List)
					    }
					    if ($Total_VDAs -eq 0)
					    {
						    $SyncHash.TB_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllVDAs.Text = "Total VDAs : $Total_VDAs    Total Servers : $Total_Servers    Total VDIs : $Total_VDIs`r`Powered Off : $Total_PoweredOff    Maintenance : $Total_Maintenance    Unregistered : $Total_Unregistered" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No VDA found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_VDAs -eq 1)
					    {
						    $AllVDAs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $AllVDAs_List_Datagrid.Add($VDAs_list)
						    $SyncHash.datagrid_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_AllVDAs.ItemsSource = $AllVDAs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllVDAs.Text = "Total VDAs : $Total_VDAs    Total Servers : $Total_Servers    Total VDIs : $Total_VDIs`r`Powered Off : $Total_PoweredOff    Maintenance : $Total_Maintenance    Unregistered : $Total_Unregistered" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_AllVDAs.text = $Simple_VDAs_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_AllVDAs_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
					    }
					    else
					    {
						    $AllVDAs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $AllVDAs_List_Datagrid.AddRange($VDAs_list)
						    $SyncHash.datagrid_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_AllVDAs.ItemsSource = $AllVDAs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TB_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.TB_AllVDAs.Text = "Total VDAs : $Total_VDAs    Total Servers : $Total_Servers    Total VDIs : $Total_VDIs`r`Powered Off : $Total_PoweredOff    Maintenance : $Total_Maintenance    Unregistered : $Total_Unregistered" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_AllVDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_AllVDAs.text = $Simple_VDAs_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_AllVDAs_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllVDAs_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Refresh_MCs
    {
	    try
	    {
		    $Farm = $datagrid_MCsList.selecteditem.Farm
		    $DDC = ($SyncHash.$Farm).DDC
		    $Name = $datagrid_MCsList.selecteditem.Name
		    $Global:SyncHash_MC_VDAs = [hashtable]::Synchronized(@{
				    Farm = $Farm
				    DDC  = $DDC
				    Name = $Name
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_MC_VDAs", $SyncHash_MC_VDAs)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $VDAs_list = Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $SyncHash_MC_VDAs.DDC | ? { $_.CatalogName -eq $SyncHash_MC_VDAs.Name } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $SyncHash_MC_VDAs.Farm } }, @{
						    n = "Type"; e = {
							    if ($_.SessionSupport -match "MultiSession") { "Server" }
							    else { "VDI" }
						    }
					    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
					    $VDAs_list = $VDAs_list | Sort-Object "Machine Name"
					    $Total_VDAs = $VDAs_list."Machine Name".count
					    $Simple_MC_VDAs_List = $VDAs_list."Machine Name"
					    if ($Total_VDAs -eq 0) { $Simple_MC_VDAs_List_String = $null }
					    else { $Simple_MC_VDAs_List_String = [string]::Join([Environment]::NewLine, $Simple_MC_VDAs_List) }
					    if ($Total_VDAs -eq 0)
					    {
						    $SyncHash.TextBox_TotalMC_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalMC_VDAs.text = "Total = $Total_VDAs" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No VDA found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_VDAs -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_MC_VDA_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $MCVDAs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $MCVDAs_List_Datagrid.Add($VDAs_list)
						    $SyncHash.datagrid_MC_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_MC_VDAs.ItemsSource = $MCVDAs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_MC_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_MC_VDAs.text = $Simple_MC_VDAs_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalMC_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalMC_VDAs.text = "Total = $Total_VDAs" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_MC_VDA_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $MCVDAs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $MCVDAs_List_Datagrid.AddRange($VDAs_list)
						    $SyncHash.datagrid_MC_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_MC_VDAs.ItemsSource = $MCVDAs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_MC_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_MC_VDAs.text = $Simple_MC_VDAs_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalMC_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalMC_VDAs.text = "Total = $Total_VDAs" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_MCs_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_MCs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Refresh_DGs
    {
	    try
	    {
		    $Farm = $datagrid_DGsList.selecteditem.Farm
		    $DDC = ($SyncHash.$Farm).DDC
		    $Name = $datagrid_DGsList.selecteditem.Name
		    $Global:SyncHash_DG_VDAs = [hashtable]::Synchronized(@{
				    Farm = $Farm
				    DDC  = $DDC
				    Name = $Name
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_DG_VDAs", $SyncHash_DG_VDAs)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $VDAs_list = Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $SyncHash_DG_VDAs.DDC | ? { $_.DesktopGroupName -eq $SyncHash_DG_VDAs.Name } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $SyncHash_DG_VDAs.Farm } }, @{
						    n = "Type"; e = {
							    if ($_.SessionSupport -match "MultiSession") { "Server" }
							    else { "VDI" }
						    }
					    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
					    $VDAs_list = $VDAs_list | Sort-Object "Machine Name"
					    $Total_VDAs = $VDAs_list."Machine Name".count
					    $Simple_DG_VDAs_List = $VDAs_list."Machine Name"
					    if ($Total_VDAs -eq 0) { $Simple_DG_VDAs_List_String = $null }
					    else { $Simple_DG_VDAs_List_String = [string]::Join([Environment]::NewLine, $Simple_DG_VDAs_List) }
					    if ($Total_VDAs -eq 0)
					    {
						    $SyncHash.TextBox_TotalDG_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_VDAs.text = "Total = $Total_VDAs" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No VDA found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_VDAs -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_DG_VDA_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $DGVDAs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $DGVDAs_List_Datagrid.Add($VDAs_list)
						    $SyncHash.datagrid_DG_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DG_VDAs.ItemsSource = $DGVDAs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_DG_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_DG_VDAs.text = $Simple_DG_VDAs_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalDG_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_VDAs.text = "Total = $Total_VDAs" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_DG_VDA_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $DGVDAs_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $DGVDAs_List_Datagrid.AddRange($VDAs_list)
						    $SyncHash.datagrid_DG_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DG_VDAs.ItemsSource = $DGVDAs_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_DG_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_DG_VDAs.text = $Simple_DG_VDAs_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalDG_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_VDAs.text = "Total = $Total_VDAs" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_DGs_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_DGs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    Function Refresh_MCSession
    {
	    try
	    {
		    $datagrid_MC_sessions.ItemsSource = $null
		    $Farm = $datagrid_MCsList.selecteditem.farm
		    $UID = $datagrid_MCsList.selecteditem.uid
		    $Name = $datagrid_MCsList.selecteditem.Name
		    $DDC = ($SyncHash.$Farm).DDC
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_MCSessions_list = [hashtable]::Synchronized(@{
				    DDC  = $DDC
				    Name = $Name
				    Farm = $Farm
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_MCSessions_list", $SyncHash_MCSessions_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $MC_Sessions_List = Get-BrokerSession -MaxRecordCount 99999 -AdminAddress $SyncHash_MCSessions_list.DDC -CatalogName $SyncHash_MCSessions_list.Name | Select-Object @{
						    n = "User"; e = {
							    if ($_.UserFullName -eq $null) { ".no data" }
							    else { $_.UserFullName }
						    }
					    }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { $SyncHash_MCSessions_list.Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
						    n = "Type"; e = {
							    if ($_.SessionSupport -match "MultiSession") { "Server" }
							    else { "VDI" }
						    }
					    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID
					    $MC_Sessions_List = $MC_Sessions_List | Sort-Object User
					    $Total_MC_Sessions = $MC_Sessions_List.User.count
					    $Simple_MC_Sessions_List = $MC_Sessions_List.User
					    if ($Total_MC_Sessions -eq 0) { $Simple_MC_Sessions_List_String = $null }
					    else { $Simple_MC_Sessions_List_String = [string]::Join([Environment]::NewLine, $Simple_MC_Sessions_List) }
					    if ($Total_MC_Sessions -eq 0)
					    {
						    $SyncHash.TextBox_TotalMC_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalMC_Sessions.text = "Total = $Total_MC_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No session found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_MC_Sessions -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_MC_Sessions_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $MCSessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $MCSessions_List_Datagrid.Add($MC_Sessions_List)
						    $SyncHash.datagrid_MC_sessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_MC_sessions.ItemsSource = $MCSessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_MC_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_MC_Sessions.text = $Simple_MC_Sessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalMC_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalMC_Sessions.text = "Total = $Total_MC_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_MC_Sessions_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $MCSessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $MCSessions_List_Datagrid.AddRange($MC_Sessions_List)
						    $SyncHash.datagrid_MC_sessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_MC_sessions.ItemsSource = $MCSessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_MC_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_MC_Sessions.text = $Simple_MC_Sessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalMC_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalMC_Sessions.text = "Total = $Total_MC_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_MCSession_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_MCSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    Function Refresh_DGSession
    {
	    try
	    {
		    $datagrid_DG_sessions.ItemsSource = $null
		    $Farm = $datagrid_DGsList.selecteditem.farm
		    $UID = $datagrid_DGsList.selecteditem.uid
		    $Name = $datagrid_DGsList.selecteditem.Name
		    $DDC = ($SyncHash.$Farm).DDC
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_DGSessions_list = [hashtable]::Synchronized(@{
				    DDC  = $DDC
				    Name = $Name
				    Farm = $Farm
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_DGSessions_list", $SyncHash_DGSessions_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $DG_Sessions_List = Get-BrokerSession -MaxRecordCount 99999 -AdminAddress $SyncHash_DGSessions_list.DDC -DesktopGroupName $SyncHash_DGSessions_list.Name | Select-Object @{
						    n = "User"; e = {
							    if ($_.UserFullName -eq $null) { ".no data" }
							    else { $_.UserFullName }
						    }
					    }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { $SyncHash_DGSessions_list.Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
						    n = "Type"; e = {
							    if ($_.SessionSupport -match "MultiSession") { "Server" }
							    else { "VDI" }
						    }
					    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID
					    $DG_Sessions_List = $DG_Sessions_List | Sort-Object User
					    $Total_DG_Sessions = $DG_Sessions_List.User.count
					    $Simple_DG_Sessions_List = $DG_Sessions_List.User
					    if ($Total_DG_Sessions -eq 0) { $Simple_DG_Sessions_List_String = $null }
					    else { $Simple_DG_Sessions_List_String = [string]::Join([Environment]::NewLine, $Simple_DG_Sessions_List) }
					    if ($Total_DG_Sessions -eq 0)
					    {
						    $SyncHash.TextBox_TotalDG_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_Sessions.text = "Total = $Total_DG_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No session found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_DG_Sessions -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_DG_Sessions_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $DGSessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $DGSessions_List_Datagrid.Add($DG_Sessions_List)
						    $SyncHash.datagrid_DG_sessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DG_sessions.ItemsSource = $DGSessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_DG_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_DG_Sessions.text = $Simple_DG_Sessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalDG_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_Sessions.text = "Total = $Total_DG_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_DG_Sessions_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $DGSessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $DGSessions_List_Datagrid.AddRange($DG_Sessions_List)
						    $SyncHash.datagrid_DG_sessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DG_sessions.ItemsSource = $DGSessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_DG_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_DG_Sessions.text = $Simple_DG_Sessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalDG_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_Sessions.text = "Total = $Total_DG_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_DGSession_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_DGSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    Function Refresh_DGPublication
    {
	    try
	    {
		    $datagrid_DG_Publications.ItemsSource = $null
		    $Farm = $datagrid_DGsList.selecteditem.farm
		    $UID = $datagrid_DGsList.selecteditem.uid
		    $Name = $datagrid_DGsList.selecteditem.Name
		    $DDC = ($SyncHash.$Farm).DDC
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_DGPublications_list = [hashtable]::Synchronized(@{
				    DDC  = $DDC
				    Name = $Name
				    Farm = $Farm
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_DGPublications_list", $SyncHash_DGPublications_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $DG_UID = Get-BrokerDesktopGroup -AdminAddress $SyncHash_DGPublications_list.DDC | Where-Object { $_.Name -eq $SyncHash_DGPublications_list.Name } | Select-Object -ExpandProperty UID
					    $DG_Applis_List = Get-BrokerApplication -MaxRecordCount 99999 -AdminAddress $SyncHash_DGPublications_list.DDC | Where-Object { $_.AssociatedDesktopGroupUids -eq $DG_UID } | Select-Object @{ n = "Application Name"; e = { $_.ApplicationName } }, @{ n = "Sessions"; e = { (Get-BrokerSession -MaxRecordCount 9999 -AdminAddress $SyncHash_DGPublications_list.DDC -ApplicationUid $_.UID).count } }, Enabled, Visible, @{ n = "Type"; e = { "Application" } }, @{ n = "Command Line Executable"; e = { $_.CommandLineExecutable } }, @{ n = "Command Line Arguments"; e = { $_.CommandLineArguments } }, @{ n = "Working Directory"; e = { $_.WorkingDirectory } }, @{ n = "Description"; e = { $_.Description } }, @{ n = "Application Groups"; e = { ForEach ($AssociatedApplicationGroupUid in $_.AssociatedApplicationGroupUids) { Get-BrokerApplicationGroup -AdminAddress $SyncHash_DGPublications_list.DDC -Uid $AssociatedApplicationGroupUid | Select-Object -ExpandProperty name } } }, UID
					    $DG_Desktops_List = Get-BrokerEntitlementPolicyRule -MaxRecordCount 9999 -AdminAddress $SyncHash_DGPublications_list.DDC -DesktopGroupUid $DG_UID | Select-Object @{ n = "Application Name"; e = { $_.Name } }, @{ n = "Sessions"; e = { (Get-BrokerSession -MaxRecordCount 9999 -AdminAddress $SyncHash_DGPublications_list.DDC -LaunchedViaPublishedName $_.PublishedName).count } }, Enabled, @{ n = "Visible"; e = { "N/A" } }, @{ n = "Type"; e = { "Desktop" } }, @{ n = "Command Line Executable"; e = { "N/A" } }, @{ n = "Command Line Arguments"; e = { "N/A" } }, @{ n = "Working Directory"; e = { "N/A" } }, @{ n = "Description"; e = { $_.Description } }, @{ n = "Application Groups"; e = { "N/A" } }, UID
					    $DG_Publications_List = @()
					    $DG_Publications_List += $DG_Applis_List
					    $DG_Publications_List += $DG_Desktops_List
					    $DG_Publications_List = $DG_Publications_List | Sort-Object "Application Name"
					    $Total_DG_Publications = $DG_Publications_List."Application Name".count
					    $Simple_DG_Publications_List = $DG_Publications_List."Application Name"
					    if ($Total_DG_Publications -eq 0) { $Simple_DG_Publications_List_String = $null }
					    else { $Simple_DG_Publications_List_String = [string]::Join([Environment]::NewLine, $Simple_DG_Publications_List) }
					    if ($Total_DG_Publications -eq 0)
					    {
						    $SyncHash.TextBox_TotalDG_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_Publications.text = "Total = $Total_DG_Publications" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No publication found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_DG_Publications -eq 1)
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_DG_Publications_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $DGPublications_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $DGPublications_List_Datagrid.Add($DG_Publications_List)
						    $SyncHash.datagrid_DG_Publications.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DG_Publications.ItemsSource = $DGPublications_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_DG_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_DG_Publications.text = $Simple_DG_Publications_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalDG_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_Publications.text = "Total = $Total_DG_Publications" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
					    else
					    {
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_DG_Publications_Full.Visibility = "Visible"
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
						    $DGPublications_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $DGPublications_List_Datagrid.AddRange($DG_Publications_List)
						    $SyncHash.datagrid_DG_Publications.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_DG_Publications.ItemsSource = $DGPublications_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_DG_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_DG_Publications.text = $Simple_DG_Publications_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalDG_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalDG_Publications.text = "Total = $Total_DG_Publications" }, [Windows.Threading.DispatcherPriority]::Normal)
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_DGPublication_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_DGPublication " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    Function Refresh_Maintenance
    {
	    try
	    {
		    $Grid_Simple_Maintenance_Registered.Visibility = "collapse"
		    $Grid_Detailled_Maintenance_Registered.Visibility = "collapse"
		    $Refresh_Maintenance.Visibility = "collapse"
		    $Refresh_Registration.Visibility = "collapse"
		    $Refresh_Maintenance_Simple.Visibility = "collapse"
		    $Refresh_Registration_Simple.Visibility = "collapse"
		    $datagrid_Maintenance_Registered.ItemsSource = $null
		    $TextBox_Servers_Maintenance_Registered.Text = ""
		    $TextBox_TotalServers_Maintenance_Registered.Text = ""
		    $Farm = $S_Maintenance.selectedItem
		    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_Maint_list = [hashtable]::Synchronized(@{
				    Farm	  = $Farm
				    DDC	      = $DDC
				    Farm_List = $SyncHash.Farm_List
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_Maint_list", $SyncHash_Maint_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $Maint_list = @()
					    $Total_Maint = @()
					    if ($SyncHash_Maint_list.Farm -eq "All Farms")
					    {
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($item in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$item).DDC
							    $Maint_list += Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $DDC | Where-Object { $_.InMaintenanceMode -eq $true } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $item } }, @{
								    n = "Type"; e = {
									    if ($_.SessionSupport -match "MultiSession") { "Server" }
									    else { "VDI" }
								    }
							    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
						    }
						    $Maint_list = $Maint_list | Sort-Object "Machine Name"
						    $Total_Maint = $Maint_list."Machine Name".count
						    $Simple_Maint_List = $Maint_list."Machine Name"
						    if ($Total_Maint -eq 0) { $Simple_Maint_List_String = $null }
						    else { $Simple_Maint_List_String = [string]::Join([Environment]::NewLine, $Simple_Maint_List) }
					    }
					    else
					    {
						    $Maint_list = Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $SyncHash_Maint_list.DDC | Where-Object { $_.InMaintenanceMode -eq $true } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $SyncHash_Maint_list.Farm } }, @{
							    n = "Type"; e = {
								    if ($_.SessionSupport -match "MultiSession") { "Server" }
								    else { "VDI" }
							    }
						    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
						    $Maint_list = $Maint_list | Sort-Object "Machine Name"
						    $Total_Maint = $Maint_list."Machine Name".count
						    $Simple_Maint_List = $Maint_list."Machine Name"
						    if ($Total_Maint -eq 0) { $Simple_Maint_List_String = $null }
						    else { $Simple_Maint_List_String = [string]::Join([Environment]::NewLine, $Simple_Maint_List) }
					    }
					    if ($Total_Maint -eq 0)
					    {
						    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Maint" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_Detailled_Maintenance_Registered.Visibility = "Visible"
								    $SyncHash.Refresh_Maintenance.Visibility = "Visible"
								    $SyncHash.Refresh_Maintenance_Simple.Visibility = "Visible"
								    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $false
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No VDA found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_Maint -eq 1)
					    {
						    $Maint_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $Maint_List_Datagrid.Add($Maint_list)
						    $SyncHash.datagrid_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_Maintenance_Registered.ItemsSource = $Maint_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_Servers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Maintenance_Registered.text = $Simple_Maint_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Maint" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_Detailled_Maintenance_Registered.Visibility = "Visible"
								    $SyncHash.Refresh_Maintenance.Visibility = "Visible"
								    $SyncHash.Refresh_Maintenance_Simple.Visibility = "Visible"
								    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $false
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
					    }
					    else
					    {
						    $Maint_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $Maint_List_Datagrid.AddRange($Maint_list)
						    $SyncHash.datagrid_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_Maintenance_Registered.ItemsSource = $Maint_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_Servers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Maintenance_Registered.text = $Simple_Maint_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Maint" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_Detailled_Maintenance_Registered.Visibility = "Visible"
								    $SyncHash.Refresh_Maintenance.Visibility = "Visible"
								    $SyncHash.Refresh_Maintenance_Simple.Visibility = "Visible"
								    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $false
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Maintenance_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Maintenance " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    Function Refresh_Registration
    {
	    try
	    {
		    $Grid_Simple_Maintenance_Registered.Visibility = "collapse"
		    $Grid_Detailled_Maintenance_Registered.Visibility = "collapse"
		    $Refresh_Maintenance.Visibility = "collapse"
		    $Refresh_Registration.Visibility = "collapse"
		    $Refresh_Maintenance_Simple.Visibility = "collapse"
		    $Refresh_Registration_Simple.Visibility = "collapse"
		    $datagrid_Maintenance_Registered.ItemsSource = $null
		    $TextBox_Servers_Maintenance_Registered.Text = ""
		    $TextBox_TotalServers_Maintenance_Registered.Text = ""
		    $Farm = $S_Registration.selectedItem
		    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    $Global:SyncHash_Regist_list = [hashtable]::Synchronized(@{
				    Farm	  = $Farm
				    DDC	      = $DDC
				    Farm_List = $SyncHash.Farm_List
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Runspace.SessionStateProxy.SetVariable("SyncHash_Regist_list", $SyncHash_Regist_list)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    $Regist_list = @()
					    $Total_Regist = @()
					    if ($SyncHash_Regist_list.Farm -eq "All Farms")
					    {
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($item in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$item).DDC
							    $Regist_list += Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $DDC | Where-Object { $_.RegistrationState -eq "Unregistered" } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $item } }, @{
								    n = "Type"; e = {
									    if ($_.SessionSupport -match "MultiSession") { "Server" }
									    else { "VDI" }
								    }
							    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
						    }
						    $Regist_list = $Regist_list | Sort-Object "Machine Name"
						    $Total_Regist = $Regist_list."Machine Name".count
						    $Simple_Regist_list = $Regist_list."Machine Name"
						    if ($Total_Regist -eq 0) { $Simple_Regist_list_String = $null }
						    else { $Simple_Regist_list_String = [string]::Join([Environment]::NewLine, $Simple_Regist_list) }
					    }
					    else
					    {
						    $Regist_list = Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $SyncHash_Regist_list.DDC | Where-Object { $_.RegistrationState -eq "Unregistered" } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $SyncHash_Regist_list.Farm } }, @{
							    n = "Type"; e = {
								    if ($_.SessionSupport -match "MultiSession") { "Server" }
								    else { "VDI" }
							    }
						    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
						    $Regist_list = $Regist_list | Sort-Object "Machine Name"
						    $Total_Regist = $Regist_list."Machine Name".count
						    $Simple_Regist_list = $Regist_list."Machine Name"
						    if ($Total_Regist -eq 0) { $Simple_Regist_list_String = $null }
						    else { $Simple_Regist_list_String = [string]::Join([Environment]::NewLine, $Simple_Regist_list) }
					    }
					    if ($Total_Regist -eq 0)
					    {
						    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Regist" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_Detailled_Maintenance_Registered.Visibility = "Visible"
								    $SyncHash.Refresh_Registration.Visibility = "Visible"
								    $SyncHash.Refresh_Registration_Simple.Visibility = "Visible"
								    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $true
								    $SyncHash.MainLayer.IsEnabled = $false
								    $SyncHash.Main_MB.Foreground = "Red"
								    $SyncHash.Main_MB.FontSize = "20"
								    $SyncHash.Main_MB.text = "No VDA found."
								    $SyncHash.Dialog_Main.IsOpen = $True
							    }, "Normal")
					    }
					    elseif ($Total_Regist -eq 1)
					    {
						    $Regist_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $Regist_List_Datagrid.Add($Regist_list)
						    $SyncHash.datagrid_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_Maintenance_Registered.ItemsSource = $Regist_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_Servers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Maintenance_Registered.text = $Simple_Regist_list_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Regist" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_Detailled_Maintenance_Registered.Visibility = "Visible"
								    $SyncHash.Refresh_Registration.Visibility = "Visible"
								    $SyncHash.Refresh_Registration_Simple.Visibility = "Visible"
								    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $true
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
					    }
					    else
					    {
						    $Regist_List_Datagrid = New-Object System.Collections.Generic.List[Object]
						    $Regist_List_Datagrid.AddRange($Regist_list)
						    $SyncHash.datagrid_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_Maintenance_Registered.ItemsSource = $Regist_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_Servers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Maintenance_Registered.text = $Simple_Regist_list_String }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Regist" }, [Windows.Threading.DispatcherPriority]::Normal)
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.Grid_Detailled_Maintenance_Registered.Visibility = "Visible"
								    $SyncHash.Refresh_Registration.Visibility = "Visible"
								    $SyncHash.Refresh_Registration_Simple.Visibility = "Visible"
								    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $true
								    $SyncHash.MainLayer.IsEnabled = $true
							    }, "Normal")
					    }
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Registration_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Registration " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Worker
    {
	    try
	    {
		    $Worker.Runspace = $Runspace
		    Register-ObjectEvent -InputObject $Worker -EventName InvocationStateChanged -Action {
			    param ([System.Management.Automation.PowerShell]$ps)
			    $state = $EventArgs.InvocationStateInfo.State
			    if ($state -in 'Completed', 'Failed')
			    {
				    $ps.EndInvoke($Worker)
				    $ps.Runspace.Dispose()
				    $ps.Dispose()
				    [GC]::Collect()
			    }
		    } | Out-Null
		    Register-ObjectEvent -InputObject $Runspace -EventName AvailabilityChanged -Action {
			    if ($($EventArgs.RunspaceAvailability) -eq 'Available')
			    {
				    $Runspace.Dispose()
				    [GC]::Collect()
			    }
		    } | Out-Null
		    $Worker.BeginInvoke()
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Worker " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Publications_collapse
    {
	    $datagrid_application_settings.Visibility = "Collapsed"
	    $datagrid_application_settings_2.Visibility = "Collapsed"
	    $datagrid_desktop_settings.Visibility = "Collapsed"
	    $datagrid_desktop_settings_2.Visibility = "Collapsed"
	    $Application_settings_Apply.Visibility = "Collapsed"
	    $Application_settings_Discard.Visibility = "Collapsed"
	    $Desktop_settings_Apply.Visibility = "Collapsed"
	    $Desktop_settings_Discard.Visibility = "Collapsed"
	    $listbox_desktop_tag.Visibility = "Collapsed"
	    $border_listbox_desktop_tag.Visibility = "Collapsed"
	    $desktop_tag.Visibility = "Collapsed"
	    $desktop_tag_remove.Visibility = "Collapsed"
	    $label_desktop_tag_list.Visibility = "Collapsed"
	    $Grid_PubliSessions_Full.Visibility = "Collapsed"
	    $Grid_PubliSessions_Simple.Visibility = "Collapsed"
	    $Grid_AllPublications_Full.Visibility = "Collapsed"
	    $Grid_AllPublications_Simple.Visibility = "Collapsed"
	    $TextBox_Servers_Publications.Visibility = "Collapsed"
	    $TextBox_TotalServers_Publications.Visibility = "Collapsed"
	    $Border_Servers_Publication.Visibility = "Collapsed"
	    $TextBox_Access_Publications.Visibility = "Collapsed"
	    $TextBox_TotalAccess_Publications.Visibility = "Collapsed"
	    $Border_Access_Publication.Visibility = "Collapsed"
    }
    function VDAs_collapse
    {
	    $datagrid_VDA_settings.Visibility = "Collapse"
	    $Enable_Maintenance_VDA.Visibility = "Collapse"
	    $Disable_Maintenance_VDA.Visibility = "Collapse"
	    $PowerOn_VDA.Visibility = "Collapse"
	    $PowerOff_VDA.Visibility = "Collapse"
	    $Refresh_VDA.Visibility = "Collapse"
	    $Grid_AllVDAs_Simple.Visibility = "Collapse"
	    $Grid_VDAs_Sessions_Full.Visibility = "Collapse"
	    $Grid_VDAs_Sessions_Simple.Visibility = "Collapse"
	    $Border_VDA_Publications.Visibility = "Collapse"
	    $TextBox_VDA_Publications.Visibility = "Collapse"
	    $TextBox_TotalVDA_Publications.Visibility = "Collapse"
	    $Border_VDA_Hotfixes.Visibility = "Collapse"
	    $datagrid_VDA_Hotfixes.Visibility = "Collapse"
	    $TextBox_TotalVDA_Hotfixes.Visibility = "Collapse"
    }
    function MCs_collapse
    {
	    $datagrid_MC_settings.Visibility = "Collapse"
	    $Grid_MC_Sessions_Full.Visibility = "Collapse"
	    $Grid_MC_Sessions_Simple.Visibility = "Collapse"
	    $Grid_MC_VDA_Full.Visibility = "Collapse"
	    $Grid_MC_VDAs_Simple.Visibility = "Collapse"
    }
    function DGs_collapse
    {
	    $datagrid_DG_settings.Visibility = "Collapse"
	    $Grid_DG.Visibility = "Collapse"
	    $Grid_DG_Desk_settings.Visibility = "Collapse"
	    $Grid_DG_Reboot_settings.Visibility = "Collapse"
	    $Grid_DG_Sessions_Full.Visibility = "Collapse"
	    $Grid_DG_Sessions_Simple.Visibility = "Collapse"
	    $Grid_DG_VDA_Full.Visibility = "Collapse"
	    $Grid_DG_VDAs_Simple.Visibility = "Collapse"
	    $Grid_DG_Publications_Full.Visibility = "Collapse"
	    $Grid_DG_Publications_Simple.Visibility = "Collapse"
    }
    function Show-Dialog_Main ($Foreground = $Main_MB.Foreground, $Text)
    {
	    try
	    {
		    $MainLayer.IsEnabled = $false
		    $Main_MB.Foreground = $Foreground
		    $Main_MB.FontSize = "20"
		    $Main_MB.text = $Text
		    $Dialog_Main.IsOpen = $True
		    $Main_MB_Close.add_Click({
				    $Dialog_Main.IsOpen = $False
				    $MainLayer.IsEnabled = $true
			    })
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Show-Dialog_Main " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Process-FarmData
    {
	    param (
		    [Parameter(Mandatory = $true)]
		    $datas
	    )
	    try
	    {
		    $MainLayer.IsEnabled = $false
		    $SpinnerOverlayLayer_Main.Visibility = "Visible"
		    if (Test-Path $global:LogoFile) { $Logo_Main.source = $global:LogoFile }
		    else { $Logo_Main.source = $DefaultLogo }
		    $Global:SyncHash = [hashtable]::Synchronized(@{
				    Form									    = $Form
				    SpinnerOverlayLayer_Main				    = $SpinnerOverlayLayer_Main
				    MainLayer								    = $MainLayer
				    datas									    = $datas
				    ListView_Main							    = $ListView_Main
				    Licenses_TB								    = $Licenses_TB
				    S_License								    = $S_License
				    Sessions_TB								    = $Sessions_TB
				    S_Sessions								    = $S_Sessions
				    VDAs_TB									    = $VDAs_TB
				    S_VDAs									    = $S_VDAs
				    Publications_TB							    = $Publications_TB
				    S_Publications							    = $S_Publications
				    Main_MB									    = $Main_MB
				    Dialog_Main								    = $Dialog_Main
				    Main_MB_Close							    = $Main_MB_Close
				    datagrid_usersList						    = $datagrid_usersList
				    Search_sessions							    = $Search_sessions
				    datagrid_UserSessions					    = $datagrid_UserSessions
				    Kill_Session							    = $Kill_Session
				    Hide_Session							    = $Hide_Session
				    Shadow_Session							    = $Shadow_Session
				    Refresh_Session							    = $Refresh_Session
				    S_Sessions_Details						    = $S_Sessions_Details
				    datagrid_AllSessions					    = $datagrid_AllSessions
				    Kill_AllSessions						    = $Kill_AllSessions
				    Hide_AllSessions						    = $Hide_AllSessions
				    Shadow_AllSessions						    = $Shadow_AllSessions
				    Refresh_AllSessions						    = $Refresh_AllSessions
				    Export_AllSessions						    = $Export_AllSessions
				    Border_AllSessions						    = $Border_AllSessions
				    TB_AllSessions							    = $TB_AllSessions
				    datagrid_publications					    = $datagrid_publications
				    Publication_settings					    = $Publication_settings
				    Publication_sessions					    = $Publication_sessions
				    Publication_servers						    = $Publication_servers
				    Publication_access						    = $Publication_access
				    datagrid_PubliSessions					    = $datagrid_PubliSessions
				    TextBox_Servers_Publications			    = $TextBox_Servers_Publications
				    TextBox_TotalServers_Publications		    = $TextBox_TotalServers_Publications
				    TextBox_Access_Publications				    = $TextBox_Access_Publications
				    TextBox_TotalAccess_Publications		    = $TextBox_TotalAccess_Publications
				    S_Publications_Details					    = $S_Publications_Details
				    datagrid_AllPublications				    = $datagrid_AllPublications
				    Disable_AllPublis						    = $Disable_AllPublis
				    Enable_AllPublis						    = $Enable_AllPublis
				    Delete_AllPublis						    = $Delete_AllPublis
				    Export_AllPublis						    = $Export_AllPublis
				    Border_AllPublis						    = $Border_AllPublis
				    TB_AllPublis							    = $TB_AllPublis
				    datagrid_VDAsList						    = $datagrid_VDAsList
				    S_VDAs_Details							    = $S_VDAs_Details
				    VDA_Registration_TB						    = $VDA_Registration_TB
				    VDA_Maintenance_TB						    = $VDA_Maintenance_TB
				    VDA_Power_TB							    = $VDA_Power_TB
				    datagrid_VDA_sessions					    = $datagrid_VDA_sessions
				    TextBox_VDA_Publications				    = $TextBox_VDA_Publications
				    TextBox_TotalVDA_Publications			    = $TextBox_TotalVDA_Publications
				    Border_VDA_Publications					    = $Border_VDA_Publications
				    datagrid_VDA_Hotfixes					    = $datagrid_VDA_Hotfixes
				    TextBox_TotalVDA_Hotfixes				    = $TextBox_TotalVDA_Hotfixes
				    Border_VDA_Hotfixes						    = $Border_VDA_Hotfixes
				    datagrid_AllVDAs						    = $datagrid_AllVDAs
				    TB_AllVDAs								    = $TB_AllVDAs
				    Border_AllVDAs							    = $Border_AllVDAs
				    datagrid_MCsList						    = $datagrid_MCsList
				    datagrid_MC_VDAs						    = $datagrid_MC_VDAs
				    datagrid_MC_sessions					    = $datagrid_MC_sessions
				    S_MCs_Details							    = $S_MCs_Details
				    TB_AllMCs								    = $TB_AllMCs
				    Border_AllMCs							    = $Border_AllMCs
				    datagrid_DGsList						    = $datagrid_DGsList
				    datagrid_DG_VDAs						    = $datagrid_DG_VDAs
				    datagrid_DG_sessions					    = $datagrid_DG_sessions
				    S_DGs_Details							    = $S_DGs_Details
				    TB_AllDGs								    = $TB_AllDGs
				    Border_AllDGs							    = $Border_AllDGs
				    S_Maintenance							    = $S_Maintenance
				    S_Registration							    = $S_Registration
				    datagrid_Maintenance_Registered			    = $datagrid_Maintenance_Registered
				    TextBox_Servers_Maintenance_Registered	    = $TextBox_Servers_Maintenance_Registered
				    TextBox_TotalServers_Maintenance_Registered = $TextBox_TotalServers_Maintenance_Registered
				    Grid_Simple_Maintenance_Registered		    = $Grid_Simple_Maintenance_Registered
				    Grid_Detailled_Maintenance_Registered	    = $Grid_Detailled_Maintenance_Registered
				    Refresh_Maintenance						    = $Refresh_Maintenance
				    Refresh_Registration					    = $Refresh_Registration
				    Enable_Maintenance_MaintRegist			    = $Enable_Maintenance_MaintRegist
				    Refresh_Maintenance_Simple				    = $Refresh_Maintenance_Simple
				    Refresh_Registration_Simple				    = $Refresh_Registration_Simple
				    TextBox_DG_Sessions						    = $TextBox_DG_Sessions
				    TextBox_TotalDG_Sessions				    = $TextBox_TotalDG_Sessions
				    Grid_DG_Sessions_Full					    = $Grid_DG_Sessions_Full
				    TextBox_DG_VDAs							    = $TextBox_DG_VDAs
				    TextBox_TotalDG_VDAs					    = $TextBox_TotalDG_VDAs
				    Grid_DG_VDA_Full						    = $Grid_DG_VDA_Full
				    TextBox_MC_Sessions						    = $TextBox_MC_Sessions
				    TextBox_TotalMC_Sessions				    = $TextBox_TotalMC_Sessions
				    Grid_MC_Sessions_Full					    = $Grid_MC_Sessions_Full
				    TextBox_MC_VDAs							    = $TextBox_MC_VDAs
				    TextBox_TotalMC_VDAs					    = $TextBox_TotalMC_VDAs
				    Grid_MC_VDA_Full						    = $Grid_MC_VDA_Full
				    TextBox_TotalVDAs_Sessions				    = $TextBox_TotalVDAs_Sessions
				    TextBox_VDAs_Sessions					    = $TextBox_VDAs_Sessions
				    Grid_VDAs_Sessions_Full					    = $Grid_VDAs_Sessions_Full
				    TextBox_AllVDAs							    = $TextBox_AllVDAs
				    Grid_AllVDAs_Full						    = $Grid_AllVDAs_Full
				    Grid_AllPublications_Full				    = $Grid_AllPublications_Full
				    TextBox_AllPublications					    = $TextBox_AllPublications
				    TextBox_TotalPubliSessions				    = $TextBox_TotalPubliSessions
				    TextBox_PubliSessions					    = $TextBox_PubliSessions
				    Grid_PubliSessions_Full					    = $Grid_PubliSessions_Full
				    TextBox_AllSessions						    = $TextBox_AllSessions
				    Grid_AllPSessions_Full					    = $Grid_AllPSessions_Full
				    TextBox_DG_Publications					    = $TextBox_DG_Publications
				    TextBox_TotalDG_Publications			    = $TextBox_TotalDG_Publications
				    Grid_DG_Publications_Full				    = $Grid_DG_Publications_Full
				    datagrid_DG_Publications				    = $datagrid_DG_Publications
			    })
		    $Runspace = [runspacefactory]::CreateRunspace()
		    $Runspace.ThreadOptions = "ReuseThread"
		    $Runspace.ApartmentState = "STA"
		    $Runspace.Open()
		    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
		    $Worker = [PowerShell]::Create().AddScript({
				    try
				    {
					    asnp Citrix*
					    ####_ListView_Main
					    $Farm_Infos = @()
					    $SyncHash.Farm_List = @()
					    $Farm_ListView = New-Object System.Collections.Generic.List[Object]
					    foreach ($data in $SyncHash.datas)
					    {
						    $DC_State = @()
						    $FarmName = $data.Farm
						    $Version = $data.Version
						    if (Test-Path Variable:\$FarmName) { Remove-Variable $FarmName }
						    $DDC = @()
						    ForEach ($DC in $data.DDC -split "`n")
						    {
							    if ((Test-Connection $DC -Count 1 -ErrorAction SilentlyContinue) -and $DC -match $env:COMPUTERNAME)
							    {
								    if ((Get-Service -Name "CitrixBrokerService").Status -eq "Running")
								    {
									    $DC_State += "OK"
									    $DDC += ,$DC
								    }
								    else
								    {
									    $DC_State += "KO"
									    $DDC += ,$DC
								    }
							    }
							    elseif ((Test-Connection $DC -Count 1 -ErrorAction SilentlyContinue) -and $DC -notmatch $env:COMPUTERNAME)
							    {
								    if (((Invoke-Command -ComputerName $DC -ScriptBlock { Get-Service -Name "CitrixBrokerService" }).Status) -eq "Running")
								    {
									    $DC_State += "OK"
									    $DDC += ,$DC
								    }
								    else
								    {
									    $DC_State += "KO"
									    $DDC += ,$DC
								    }
							    }
							    else
							    {
								    $DC_State += "KO"
								    $DDC += ,$DC
							    }
						    }
						    ForEach ($DC in $data.DDC -split "`n")
						    {
							    if ((Test-Connection $DC -Count 1 -ErrorAction SilentlyContinue) -and $DC -match $env:COMPUTERNAME)
							    {
								    if ((Get-Service -Name "CitrixBrokerService").Status -eq "Running")
								    {
									    $Test_DB = (Test-BrokerDBConnection (Get-BrokerDBConnection -AdminAddress $DC)).ServiceStatus
									    if ($Test_DB -eq "OK")
									    {
										    $VariableName = "${FarmName}"
										    $VariableValue = @{ Farm = $FarmName; DDC = $DC }
										    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
										    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
										    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
										    New-Variable -Name $VariableName -Value $VariableValue
										    $SyncHash.Add($VariableName, $VariableValue)
										    $SyncHash.Farm_List = $SyncHash.Farm_List + $VariableName
										    break
									    }
								    }
								    else { $Test_DB = "KO" }
							    }
							    elseif ((Test-Connection $DC -Count 1 -ErrorAction SilentlyContinue) -and $DC -notmatch $env:COMPUTERNAME)
							    {
								    if (((Invoke-Command -ComputerName $DC -ScriptBlock { Get-Service -Name "CitrixBrokerService" }).Status) -eq "Running")
								    {
									    $Test_DB = (Test-BrokerDBConnection (Get-BrokerDBConnection -AdminAddress $DC)).ServiceStatus
									    if ($Test_DB -eq "OK")
									    {
										    $VariableName = "${FarmName}"
										    $VariableValue = @{ Farm = $FarmName; DDC = $DC }
										    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
										    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
										    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
										    New-Variable -Name $VariableName -Value $VariableValue
										    $SyncHash.Add($VariableName, $VariableValue)
										    $SyncHash.Farm_List = $SyncHash.Farm_List + $VariableName
										    break
									    }
								    }
								    else { $Test_DB = "KO" }
							    }
							    else { $Test_DB = "KO" }
						    }
						    $Farm_infos = New-Object -TypeName PSObject -Property @{ "Farm" = $data.Farm; "Version" = $data.Version; "DDC" = $DDC; "DDC_State" = $DC_State; "Farm_State" = $Test_DB }
						    $Farm_ListView.Add($Farm_Infos)
					    }
					    ####_License
					    $DDC_License = (Get-Variable -Name $SyncHash.Farm_List[0] -ValueOnly).DDC
					    $License_server = (Get-BrokerSite -AdminAddress $DDC_License).LicenseServerName
					    $CertHash = (Get-LicCertificate -AdminAddress $License_server).CertHash
					    $Lic_Inventory = Get-LicInventory -AdminAddress $License_server -CertHash $CertHash | Sort-Object -Descending LicenseProductName
					    $Lic_List = @()
					    $SyncHash.Lic_Var = @()
					    foreach ($Type in $Lic_Inventory)
					    {
						    $LicenseProductName = $Type.LicenseProductName
						    $VariableName = "${LicenseProductName}"
						    $LocalizedLicenseProductName = $Type.LocalizedLicenseProductName
						    $LicenseEdition = $Type.LicenseEdition
						    $LicenseSubscriptionAdvantageDate = ($Type.LicenseSubscriptionAdvantageDate).ToString('yyyy.MMdd')
                            $LicenseExpirationDate = $Type.LicenseExpirationDate
                            $LocalizedLicenseModel = $Type.LocalizedLicenseModel
						    [int]$InUseCount = $Type.LicensesInUse
						    [int]$Count = $Type.LicensesAvailable
						    [int]$Left = $Count - $InUseCount
						    $LicenseModel = $Type.LicenseModel
						    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
						    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
						    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
						    $VariableValue = @{ ProductName = $LocalizedLicenseProductName; LicenseEdition = if ($LicenseEdition.Length -eq 0) { "N/A" } else {$LicenseEdition}; SubscriptionAdvantageDate = $LicenseSubscriptionAdvantageDate; LicenseExpirationDate = $LicenseExpirationDate; LocalizedLicenseModel = $LocalizedLicenseModel; InUseCount = $InUseCount; Count = $Count; LicenseModel = if ($LicenseModel.Length -eq 0) { "N/A" } else {$LicenseModel}; Left = $Left }
						    New-Variable -Name $VariableName -Value $VariableValue
						    $SyncHash.Add($VariableName, $VariableValue)
						    $Lic_List += $Type.LocalizedLicenseProductName
						    $SyncHash.Lic_Var += $VariableName
					    }
					    foreach ($item in $Lic_List) { $SyncHash.S_License.Dispatcher.Invoke([Action]{ $SyncHash.S_License.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    ####_Sessions
					    $SyncHash.Total_Sessions_All = $null
					    $SyncHash.Active_Sessions_All = $null
					    $SyncHash.Disconnected_Sessions_All = $null
					    $SyncHash.Connected_Sessions_All = $null
					    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
					    foreach ($Farm in $SyncHash.Farm_List)
					    {
						    $DDC = ($SyncHash.$Farm).DDC
						    $Sessions = Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $DDC
						    $Total_Sessions = $Sessions.count
						    $SyncHash.Total_Sessions_All += $Total_Sessions
						    $Active_Sessions = ($Sessions | Where-Object { $_.SessionState -eq "Active" }).count
						    $SyncHash.Active_Sessions_All += $Active_Sessions
						    $Disconnected_Sessions = ($Sessions | Where-Object { $_.SessionState -eq "Disconnected" }).count
						    $SyncHash.Disconnected_Sessions_All += $Disconnected_Sessions
						    $Connected_Sessions = ($Sessions | Where-Object { $_.SessionState -eq "Connected" }).count
						    $SyncHash.Connected_Sessions_All += $Connected_Sessions
						    $VariableName = "Sessions_" + "${Farm}"
						    $VariableValue = @{ Total_Sessions = $Total_Sessions; Active_Sessions = $Active_Sessions; Disconnected_Sessions = $Disconnected_Sessions; Connected_Sessions = $Connected_Sessions }
						    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
						    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
						    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
						    New-Variable -Name $VariableName -Value $VariableValue
						    $SyncHash.Add($VariableName, $VariableValue)
					    }
					    $VariableName = "Sessions_All farms"
					    $VariableValue = @{ Total_Sessions = $SyncHash.Total_Sessions_All; Active_Sessions = $SyncHash.Active_Sessions_All; Disconnected_Sessions = $SyncHash.Disconnected_Sessions_All; Connected_Sessions = $SyncHash.Connected_Sessions_All }
					    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
					    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
					    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
					    New-Variable -Name $VariableName -Value $VariableValue
					    $SyncHash.Add($VariableName, $VariableValue)
					    $SyncHash.Farm_List = $SyncHash.Farm_List | Sort-Object
					    if ($SyncHash.Farm_List.count -ne 1 -and $SyncHash.Farm_List -notcontains "All farms")
					    {
						    $Array = @("All farms") + $SyncHash.Farm_List
						    $SyncHash.Farm_List = $Array
					    }
					    else { $SyncHash.Farm_List = ,$SyncHash.Farm_List }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.S_Sessions.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_Sessions_Details.Dispatcher.Invoke([Action]{ $SyncHash.S_Sessions_Details.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    ###_VDAs
					    $SyncHash.Total_VDAs_All = $null
					    $SyncHash.Total_Servers_All = $null
					    $SyncHash.Total_VDIs_All = $null
					    $SyncHash.PoweredOff_All = $null
					    $SyncHash.Maintenance_All = $null
					    $SyncHash.Unregistered_All = $null
					    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
					    foreach ($Farm in $SyncHash.Farm_List)
					    {
						    $DDC = ($SyncHash.$Farm).DDC
						    $VDAs = Get-BrokerDesktop -MaxRecordCount 999999 -AdminAddress $DDC
						    $Total_VDAs = $VDAs.count
						    $SyncHash.Total_VDAs_All += $Total_VDAs
						    $Total_Servers = ($VDAs | Where-Object { $_.DesktopKind -eq "Shared" }).count
						    $SyncHash.Total_Servers_All += $Total_Servers
						    $Total_VDIs = ($VDAs | Where-Object { $_.DesktopKind -eq "Private" }).count
						    $SyncHash.Total_VDIs_All += $Total_VDIs
						    $PoweredOff = ($VDAs | Where-Object { $_.PowerState -eq "Off" }).count
						    $SyncHash.PoweredOff_All += $PoweredOff
						    $Maintenance = ($VDAs | Where-Object { $_.InMaintenanceMode -eq $true }).count
						    $SyncHash.Maintenance_All += $Maintenance
						    $Unregistered = ($VDAs | Where-Object { $_.RegistrationState -eq "Unregistered" }).count
						    $SyncHash.Unregistered_All += $Unregistered
						    $VariableName = "VDAs_" + "${Farm}"
						    $VariableValue = @{ Total_VDAs = $Total_VDAs; Total_Servers = $Total_Servers; Total_VDIs = $Total_VDIs; PoweredOff = $PoweredOff; Maintenance = $Maintenance; Unregistered = $Unregistered }
						    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
						    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
						    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
						    New-Variable -Name $VariableName -Value $VariableValue
						    $SyncHash.Add($VariableName, $VariableValue)
					    }
					    $VariableName = "VDAs_All farms"
					    $VariableValue = @{ Total_VDAs = $SyncHash.Total_VDAs_All; Total_Servers = $SyncHash.Total_Servers_All; Total_VDIs = $SyncHash.Total_VDIs_All; PoweredOff = $SyncHash.PoweredOff_All; Maintenance = $SyncHash.Maintenance_All; Unregistered = $SyncHash.Unregistered_All }
					    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
					    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
					    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
					    New-Variable -Name $VariableName -Value $VariableValue
					    $SyncHash.Add($VariableName, $VariableValue)
					    $SyncHash.Farm_List = $SyncHash.Farm_List | Sort-Object
					    if ($SyncHash.Farm_List.count -ne 1 -and $SyncHash.Farm_List -notcontains "All farms")
					    {
						    $Array = @("All farms") + $SyncHash.Farm_List
						    $SyncHash.Farm_List = $Array
					    }
					    else { $SyncHash.Farm_List = ,$SyncHash.Farm_List }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.S_VDAs.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_VDAs_Details.Dispatcher.Invoke([Action]{ $SyncHash.S_VDAs_Details.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_MCs_Details.Dispatcher.Invoke([Action]{ $SyncHash.S_MCs_Details.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_DGs_Details.Dispatcher.Invoke([Action]{ $SyncHash.S_DGs_Details.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_Maintenance.Dispatcher.Invoke([Action]{ $SyncHash.S_Maintenance.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_Registration.Dispatcher.Invoke([Action]{ $SyncHash.S_Registration.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    ###_Publications
					    $SyncHash.Total_Publications_All = $null
					    $SyncHash.Total_Applications_All = $null
					    $SyncHash.Total_Desktops_All = $null
					    $SyncHash.Publications_Disabled_All = $null
					    $SyncHash.Publications_Hidden_All = $null
					    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
					    foreach ($Farm in $SyncHash.Farm_List)
					    {
						    $DDC = ($SyncHash.$Farm).DDC
						    $Applications = Get-BrokerApplication -MaxRecordCount 999999 -AdminAddress $DDC
						    $Desktops = Get-BrokerEntitlementPolicyRule -MaxRecordCount 999999 -AdminAddress $DDC
						    $Total_Applications = $Applications.count
						    $SyncHash.Total_Applications_All += $Total_Applications
						    $Total_Desktops = $Desktops.count
						    $SyncHash.Total_Desktops_All += $Total_Desktops
						    $Total_Publications = $Total_Applications + $Total_Desktops
						    $SyncHash.Total_Publications_All += $Total_Publications
						    $Publications_Disabled = ($Applications | Where-Object { $_.Enabled -eq $False }).count
						    $SyncHash.Publications_Disabled_All += $Publications_Disabled
						    $Publications_Hidden = ($Applications | Where-Object { $_.Visible -eq $False }).count
						    $SyncHash.Publications_Hidden_All += $Publications_Hidden
						    $VariableName = "Publications_" + "${Farm}"
						    $VariableValue = @{ Total_Publications = $Total_Publications; Total_Applications = $Total_Applications; Total_Desktops = $Total_Desktops; Publications_Disabled = $Publications_Disabled; Publications_Hidden = $Publications_Hidden }
						    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
						    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
						    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
						    New-Variable -Name $VariableName -Value $VariableValue
						    $SyncHash.Add($VariableName, $VariableValue)
					    }
					    $VariableName = "Publications_All farms"
					    $VariableValue = @{ Total_Publications = $SyncHash.Total_Publications_All; Total_Applications = $SyncHash.Total_Applications_All; Total_Desktops = $SyncHash.Total_Desktops_All; Publications_Disabled = $SyncHash.Publications_Disabled_All; Publications_Hidden = $SyncHash.Publications_Hidden_All }
					    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
					    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
					    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
					    New-Variable -Name $VariableName -Value $VariableValue
					    $SyncHash.Add($VariableName, $VariableValue)
					    $SyncHash.Farm_List = $SyncHash.Farm_List | Sort-Object
					    if ($SyncHash.Farm_List.count -ne 1 -and $SyncHash.Farm_List -notcontains "All farms")
					    {
						    $Array = @("All farms") + $SyncHash.Farm_List
						    $SyncHash.Farm_List = $Array
					    }
					    else { $SyncHash.Farm_List = ,$SyncHash.Farm_List }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_Publications.Dispatcher.Invoke([Action]{ $SyncHash.S_Publications.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_Publications_Details.Dispatcher.Invoke([Action]{ $SyncHash.S_Publications_Details.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
					    ###_Publications_End
					    $SyncHash.Form.Dispatcher.Invoke([action]{
							    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
							    $SyncHash.MainLayer.IsEnabled = $true
							    $SyncHash.ListView_Main.ItemsSource = $Farm_ListView
							    $SyncHash.S_License.SelectedItem = $Lic_List[0]
							    $SyncHash.S_Sessions.SelectedItem = $SyncHash.S_Sessions.Items[0]
							    $SyncHash.S_VDAs.SelectedItem = $SyncHash.S_VDAs.Items[0]
							    $SyncHash.S_Publications.SelectedItem = $SyncHash.S_Publications.Items[0]
						    }, "Normal")
				    }
				    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Process-FarmData_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
			    })
		    Worker
		    $FarmColumnIndex = -1
		    for ($i = 0; $i -lt $ListView_Main.View.Columns.Count; $i++)
		    {
			    if ($ListView_Main.View.Columns[$i].Header -eq "Farm")
			    {
				    $FarmColumnIndex = $i
				    break
			    }
		    }
		    if ($FarmColumnIndex -ne -1)
		    {
			    $ListView_Main.Items.SortDescriptions.Clear()
			    $ListView_Main.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription "Farm", "Ascending"))
			    $ListView_Main.Items.Refresh()
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Process-FarmData " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Publication_settings
    {
	    try
	    {
		    if ($datagrid_publications.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
		    else
		    {
			    $Farm = $datagrid_publications.selecteditem.Farm
			    $DDC = ($SyncHash.$Farm).DDC
			    $UID = $datagrid_publications.selecteditem.UID
			    $Type = $datagrid_publications.selecteditem.Type
			    $Name = $datagrid_publications.selecteditem.Name
			    if ($Type -eq "Application")
			    {
				    $App = Get-BrokerApplication -AdminAddress $DDC -Uid $UID
				    $App_PublishedName = $App.PublishedName
				    $App_ApplicationName = $App.ApplicationName
				    $App_BrowserName = $App.BrowserName
				    $App_Name = $App.Name
				    $App_CommandLineExecutable = $App.CommandLineExecutable
				    $App_CommandLineArguments = $App.CommandLineArguments
				    $App_Description = $App.Description
				    $App_WorkingDirectory = $App.WorkingDirectory
				    $App_AdminFolderName = $App.AdminFolderName
				    $App_ClientFolder = $App.ClientFolder
				    $App_Enabled = $App.Enabled
				    $App_Visible = $App.Visible
				    $App_RestrictToTag = $App.RestrictToTag
				    $App_DesktopGroupUid = $App.DesktopGroupUid
				    $App_IncludedUsers = $App.IncludedUsers.name
				    $App_UUID = $App.UUID
				    $DG = Get-BrokerDesktopGroup -AdminAddress $DDC | Select-Object name, uid
				    $AppsG = Get-BrokerApplicationGroup -AdminAddress $DDC | Select-Object name, uid
				    $App_DG = @()
				    $App_G = @()
				    Foreach ($DGi in $DG)
				    {
					    Foreach ($App_DG_UID in $App.AssociatedDesktopGroupUids)
					    {
						    if ($App_DG_UID -eq $DGi.uid) { $App_DG += $DGi.Name }
					    }
				    }
				    $App_DG = $App_DG | Out-String
				    Foreach ($AppG in $AppsG)
				    {
					    Foreach ($AppG_UID in $App.AssociatedApplicationGroupUids)
					    {
						    if ($AppG_UID -eq $AppG.uid) { $App_G += $AppG.Name }
					    }
				    }
				    $App_G = $App_G | Out-String
				    $datagrid_application_settings.Visibility = "Visible"
				    $datagrid_application_settings_2.Visibility = "Visible"
				    $Application_settings_Apply.Visibility = "Visible"
				    $Application_settings_Discard.Visibility = "Visible"
				    $datagrid_application_settings.ItemsSource = @(
					    [PSCustomObject]@{ Column1Header = "Farm"; Column2Data = $Farm; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Name"; Column2Data = $App_Name; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Studio folder name"; Column2Data = $App_AdminFolderName; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "PublishedName / Application name (for users)"; Column2Data = $App_PublishedName; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "ApplicationName / Application name (for administrators)"; Column2Data = $App_ApplicationName; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Browser Name"; Column2Data = $App_BrowserName; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Path to the executable"; Column2Data = $App_CommandLineExecutable; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Command line argument"; Column2Data = $App_CommandLineArguments; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Description and keywords"; Column2Data = $App_Description; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Working Directory"; Column2Data = $App_WorkingDirectory; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Client Folder"; Column2Data = $App_ClientFolder; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Delivery Groups"; Column2Data = $App_DG; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Application Groups"; Column2Data = $App_G; IsReadOnly = $true }
				    )
				    $datagrid_application_settings_2.ItemsSource = @(
					    [PSCustomObject]@{ Column1Header = "Enabled"; Column2Data = [System.Collections.ObjectModel.ObservableCollection[object]]@($true, $false); Column2SelectedValue = $App_Enabled; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Visible"; Column2Data = [System.Collections.ObjectModel.ObservableCollection[object]]@($true, $false); Column2SelectedValue = $App_Visible; IsReadOnly = $false }
				    )
			    }
			    else
			    {
				    $App = Get-BrokerEntitlementPolicyRule -AdminAddress $DDC -Uid $UID -ErrorAction SilentlyContinue
				    $App_PublishedName = $App.PublishedName
				    $App_BrowserName = $App.BrowserName
				    $App_Name = $App.Name
				    $App_Description = $App.Description
				    $App_Enabled = $App.Enabled
				    $App_RestrictToTag = $App.RestrictToTag
				    $App_IncludedUsers = $App.IncludedUsers.name
				    $App_DesktopGroupUid = $App.DesktopGroupUid
				    $DG = Get-BrokerDesktopGroup -AdminAddress $DDC | Select-Object name, uid
				    $App_DesktopGroup = $DG | ? { $_.uid -eq $App_DesktopGroupUid } | Select-Object -ExpandProperty Name
				    $datagrid_desktop_settings.Visibility = "Visible"
				    $datagrid_desktop_settings_2.Visibility = "Visible"
				    $Desktop_settings_Apply.Visibility = "Visible"
				    $Desktop_settings_Discard.Visibility = "Visible"
				    $listbox_desktop_tag.Visibility = "Visible"
				    $border_listbox_desktop_tag.Visibility = "Visible"
				    $desktop_tag.Visibility = "Visible"
				    $desktop_tag_remove.Visibility = "Visible"
				    $label_desktop_tag_list.Visibility = "Visible"
				    $datagrid_desktop_settings.ItemsSource = @(
					    [PSCustomObject]@{ Column1Header = "Farm"; Column2Data = $Farm; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "DesktopGroup"; Column2Data = $App_DesktopGroup; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Browser name"; Column2Data = $App_BrowserName; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Name"; Column2Data = $App_Name; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "PublishedName / Display name"; Column2Data = $App_PublishedName; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Description"; Column2Data = $App_Description; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Restrict To Tag"; Column2Data = $App_RestrictToTag; IsReadOnly = $True }
				    )
				    $datagrid_desktop_settings_2.ItemsSource = @(
					    [PSCustomObject]@{ Column1Header = "Enabled"; Column2Data = [System.Collections.ObjectModel.ObservableCollection[object]]@($true, $false); Column2SelectedValue = $App_Enabled; IsReadOnly = $false }
				    )
				    $TAG_List = @()
				    $TAG_List = Get-BrokerTag -AdminAddress $DDC | Select-Object -ExpandProperty Name
				    foreach ($tag in $TAG_List) { $listbox_desktop_tag.Items.Add($tag) }
			    }
		    }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publication_settings " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Main_MB_Confirm
    {
	    try
	    {
		    $Dialog_Main_Confirm.IsOpen = $False
		    $MainLayer.IsEnabled = $true
		    foreach ($VDA in $VDAs)
		    {
			    $UID = $VDA.Uid
			    $DDC = $SyncHash.($VDA.Farm).DDC
			    $MachineName = $VDA."Machine Name"
			    if ($global:Action -eq "Enable_Maintenance_AllVDAs") { Set-BrokerMachineMaintenanceMode -AdminAddress $DDC -InputObject $UID $true }
			    elseif ($global:Action -eq "Disble_Maintenance_AllVDAs") { Set-BrokerMachineMaintenanceMode -AdminAddress $DDC -InputObject $UID $false }
			    elseif ($global:Action -eq "PowerOn_AllVDAs") { New-BrokerHostingPowerAction -AdminAddress $DDC -MachineName $MachineName -Action TurnOn }
			    elseif ($global:Action -eq "PowerOff_AllVDAs") { New-BrokerHostingPowerAction -AdminAddress $DDC -MachineName $MachineName -Action TurnOff }
		    }
		    if ($global:Action -eq "Enable_Maintenance_AllVDAs") { Show-Dialog_Main -Foreground "Blue" -Text "Maintenance enabled for the selected VDAs.`r`nPlease Refresh" }
		    elseif ($global:Action -eq "Disble_Maintenance_AllVDAs") { Show-Dialog_Main -Foreground "Blue" -Text "Maintenance disabled for the selected VDAs.`r`nPlease Refresh" }
		    elseif ($global:Action -eq "PowerOn_AllVDAs") { Show-Dialog_Main -Foreground "Blue" -Text "Selected VDAs powered on.`r`nPlease Refresh" }
		    elseif ($global:Action -eq "PowerOff_AllVDAs") { Show-Dialog_Main -Foreground "Blue" -Text "Selected VDAs powered off.`r`nPlease Refresh" }
	    }
	    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Main_MB_Confirm " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    }
    function Main_MB_Cancel
    {
	    $Dialog_Main_Confirm.IsOpen = $False
	    $MainLayer.IsEnabled = $true
    }

    Hide-Console
    #############
    ### Start_XML
    #############
    try
    {
	    $XamlMainWindow = LoadXml("Configuration\XAML\Xd_Tool.xaml")
	    $Form = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XamlMainWindow))
	    $XamlMainWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{ Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) }
	
	    $Check_conf = LoadXml("Configuration\XAML\Chek_conf.xaml")
	    $Form_Check_conf = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Check_conf))
	    $Check_conf.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{ Set-Variable -Name ($_.Name) -Value $Form_Check_conf.FindName($_.Name) }
	    $Version.content = "V2.0.6"
	    $TB_About_XD_Tool.text = "This tool can't replace Studio of course, but it's not the purpose neither today or tomorow.`r`nActually, it ease the daily work of Citrix administrators.`r`nAll farms data are accessible in the same window.`r`nIt gives a quick access to informations.`r`nIt is a scalable tool.`r`nIt is written in WPF with Material Design UI, the code behind is PowerShell.`r`n`r`nThe configuration file, log file and the exports are in %LOCALAPPDATA%\XD_Tool"
	    $TB_Thanks.text = "Warren Frame for his work on PSExcel module :`r`nhttps://github.com/RamblingCookieMonster/PSExcel`r`nJames Willock for his great Material Design In XAML Toolkit:`r`nhttp://materialdesigninxaml.net`r`nJérôme Bezet-Torres for his Material Design theme Manager :`r`nhttps://jm2k69.github.io/2019/07/Material-Design-theme-manager.html`r`nAvi Coren for the spinner :`r`nhttps://www.materialdesignps.com/post/how-to-add-a-spinner`r`nAnd, of course, Google for the Material Design."
	    $Author.content = "Ramzi Mahdaoui"
	    $Contact.content = "ramzi.nbr@gmail.com"
	    $Download_link.content = "https://drive.google.com/drive/folders/1TC0kpflmsjW9AuvpJ11uZr9pjpKzaUst?usp=sharing"
    }
    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_XML " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
    ###########
    ### End_XML
    ###########
    #############################
    ### Start_Check_Configuration
    #############################
    ### Check_Configuration // Prerequisites
    $Prerequisites.add_Click({
		    try
		    {
			    $Chk_Conf_MB.Foreground = "BLue"
			    $Chk_Conf_MB.FontSize = "20"
			    $Chk_Conf_MB.text = "CVAD PowerShell SDK installed`r`nAdministrator rights on farms`r`nWinRM enabled on DDCs`r`n`r`nErrors log file : $env:LOCALAPPDATA\XD_Tool\Log_Errors.txt"
			    $Dialog_Chk_Conf.IsOpen = $True
			    $Chk_Conf_MB_Close.add_Click({ $Dialog_Chk_Conf.IsOpen = $False })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Prerequisites " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    ### Check_Configuration // Load
    $Load_Conf.add_Click({
		    try
		    {
			    $dialog = New-Object System.Windows.Forms.OpenFileDialog
			    $dialog.Multiselect = $false
			    $dialog.Filter = "XML files (config.xml)|config.xml"
			    $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
			    $result = $dialog.ShowDialog()
			    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
			    {
				    $SelectFile = $dialog.FileName
				    if ($SelectFile -eq $ConfigFile)
				    {
					    $Chk_Conf_MB.Foreground = "Red"
					    $Chk_Conf_MB.FontSize = "20"
					    $Chk_Conf_MB.text = "This 'config.xml' is currently loaded.`r`nPlease select another."
					    $Dialog_Chk_Conf.IsOpen = $True
					    $Chk_Conf_MB_Close.add_Click({ $Dialog_Chk_Conf.IsOpen = $False })
				    }
				    elseif (!((Get-Content $SelectFile) -match 'http://schemas.microsoft.com/powershell/2004/04'))
				    {
					    $Chk_Conf_MB.Foreground = "Red"
					    $Chk_Conf_MB.FontSize = "20"
					    $Chk_Conf_MB.text = "'config.xml' file not valid.`r`nPlease select a valid file or create a new one."
					    $Dialog_Chk_Conf.IsOpen = $True
					    $Chk_Conf_MB_Close.add_Click({ $Dialog_Chk_Conf.IsOpen = $False })
				    }
				    else
				    {
					    $objects = Import-Clixml -Path $SelectFile
					    if ($objects.Farm.Count -eq 0 -and $objects.Version.Count -eq 0 -and $objects.DDC.Count -eq 0)
					    {
						    $Chk_Conf_MB.Foreground = "Red"
						    $Chk_Conf_MB.FontSize = "20"
						    $Chk_Conf_MB.text = "'config.xml' file not valid.`r`nplease select a valid file or create a new one."
						    $Dialog_Chk_Conf.IsOpen = $True
						    $Chk_Conf_MB_Close.add_Click({ $Dialog_Chk_Conf.IsOpen = $False })
					    }
					    Else
					    {
						    $Form_Check_conf.Hide()
						    if (Test-Path $ConfigFile) { Remove-Item -Path $ConfigFile -Force }
						    Copy-Item $SelectFile $ConfigFile -Force
						    $datas = Import-Clixml -Path $SelectFile
						    if (Test-Path variable:\SyncHash.Farm_List) { $SyncHash.Farm_List = @() }
						    $S_Sessions.Items.Clear()
						    $S_VDAs.Items.Clear()
						    $S_Publications.Items.Clear()
						    $S_Sessions_Details.Items.Clear()
						    $S_Publications_Details.Items.Clear()
						    $S_VDAs_Details.Items.Clear()
						    $S_MCs_Details.Items.Clear()
						    $S_DGs_Details.Items.Clear()
						    $S_Maintenance.Items.Clear()
						    $S_Registration.Items.Clear()
						    $Tab_Control.SelectedIndex = 0
						    $Load_TB.Text = "Loading configuration file"
						    Process-FarmData -datas $datas
						    $Form.Add_Closing({
								    $Process = Get-Process XD_Tool -ErrorAction SilentlyContinue
								    if ($Process) { $Process | Stop-Process -Force }
							    })
						    $Form.ShowDialog() | Out-Null
					    }
				    }
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Load_Conf " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    ####################################
    ### Start_Check_Configuration // New
    ####################################                   
    $New_Conf.add_Click({
		    try
		    {
			    $Configuration = LoadXml("Configuration\XAML\Configuration.xaml")
			    $Form_Configuration = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Configuration))
			    $Configuration.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{ Set-Variable -Name ($_.Name) -Value $Form_Configuration.FindName($_.Name) }
			    $DDC_State.source = $Grey
			    $DDC_TB.text = ""
			    $Logo_Conf.Source = $null
			    if ($ListView_Conf.ItemsSource.count -ne 0)
			    {
				    $ListView_Conf.ItemsSource.Clear()
				    $ListView_Conf.ItemsSource = $null
				    $ListView_Conf.ItemsSource = @()
			    }
			    $Form_Check_conf.Hide()
			    ##############################
			    ### Start_Configuration // New
			    ##############################
			    ### Configuration // Test_DDC
			    $DDC_TB.Add_KeyDown({
					    param ($sender,
						    $e)
					    if ($e.Key -eq [System.Windows.Input.Key]::Enter) { Test_DDC }
				    })
			    $Test_DDC.add_Click({ Test_DDC })
			    ### Configuration // Add_DDC
			    $Add_DDC.add_Click({
					    try
					    {
						    $Farm_Infos = @()
						    $Farm_List_Conf = @()
						    $DDC = @()
						    if ($DDC_State.source -match "Grey")
						    {
							    $ApplicationLayer.IsEnabled = $false
							    $DDC_MB.Foreground = "Red"
							    $DDC_MB.FontSize = "20"
							    $DDC_MB.text = "Please test and valid a DDC first."
							    $Dialog_DDC.IsOpen = $True
							    $DDC_MB_Close.add_Click({
									    $Dialog_DDC.IsOpen = $False
									    $ApplicationLayer.IsEnabled = $true
								    })
						    }
						    Elseif ($DDC_State.source -match "Red")
						    {
							    $ApplicationLayer.IsEnabled = $false
							    $DDC_MB.Foreground = "Red"
							    $DDC_MB.FontSize = "20"
							    $DDC_MB.text = "Please enter a valid DDC."
							    $Dialog_DDC.IsOpen = $True
							    $DDC_MB_Close.add_Click({
									    $Dialog_DDC.IsOpen = $False
									    $ApplicationLayer.IsEnabled = $true
								    })
						    }
						    Elseif ($DDC_State.source -match "Green")
						    {
							    $SpinnerOverlayLayer.Visibility = "Visible"
							    $Main_Load_TB.text = "Farm adding in progress"
							    $Global:SyncHash_Conf = [hashtable]::Synchronized(@{
									    Form_Configuration  = $Form_Configuration
									    SpinnerOverlayLayer = $SpinnerOverlayLayer
									    ApplicationLayer    = $ApplicationLayer
									    DDC_MB			    = $DDC_MB
									    Dialog_DDC		    = $Dialog_DDC
									    DDC_State		    = $DDC_State
									    Green			    = $Green
									    Red				    = $Red
									    Grey			    = $Grey
									    DDC_TB			    = $DDC_TB.text
									    ListView_Conf	    = $ListView_Conf
								    })
							    $Runspace = [runspacefactory]::CreateRunspace()
							    $Runspace.ThreadOptions = "ReuseThread"
							    $Runspace.ApartmentState = "STA"
							    $Runspace.Open()
							    $Runspace.SessionStateProxy.SetVariable("SyncHash_Conf", $SyncHash_Conf)
							    $Worker = [PowerShell]::Create().AddScript({
									    try
									    {
										    asnp Citrix*
										    $Farm = (Get-BrokerSite -AdminAddress $SyncHash_Conf.DDC_TB).Name
										    if ($SyncHash_Conf.ListView_Conf.itemssource.farm -contains $Farm)
										    {
											    $SyncHash_Conf.Form_Configuration.Dispatcher.Invoke([action]{
													    $SyncHash_Conf.ApplicationLayer.IsEnabled = $false
													    $SyncHash_Conf.SpinnerOverlayLayer.Visibility = "Collapsed"
													    $SyncHash_Conf.DDC_MB.Foreground = "Red"
													    $SyncHash_Conf.DDC_MB.FontSize = "20"
													    $SyncHash_Conf.DDC_MB.text = "Farm $Farm aleady added.`r`nPlease enter a DDC from a new farm."
													    $SyncHash_Conf.Dialog_DDC.IsOpen = $True
													    $SyncHash_Conf.DDC_State.source = $SyncHash_Conf.Grey
												    }, "Normal")
										    }
										    else
										    {
											    $Version = (Get-BrokerController -AdminAddress $SyncHash_Conf.DDC_TB | Select-Object -First 1).ControllerVersion
											    $DDC = Get-BrokerController -AdminAddress $SyncHash_Conf.DDC_TB | Select-Object -exp DNSName
											    $DDC_State = @()
											    Foreach ($item in $DDC)
											    {
												    if ((Test-Connection $item -Count 1 -ErrorAction SilentlyContinue) -and $item -match $env:COMPUTERNAME)
												    {
													    if (((Get-Service -Name "CitrixBrokerService").Status) -eq "Running")
													    {
														    $DDC_State += "OK"
														    $DC += ,$item
													    }
													    else
													    {
														    $DDC_State += "KO"
														    $DC += ,$item
													    }
												    }
												    elseif ((Test-Connection $item -Count 1 -ErrorAction SilentlyContinue) -and $item -notmatch $env:COMPUTERNAME)
												    {
													    if (((Invoke-Command -ComputerName $item -ScriptBlock { Get-Service -Name "CitrixBrokerService" }).Status) -eq "Running")
													    {
														    $DDC_State += "OK"
														    $DC += ,$item
														
													    }
													    else
													    {
														    $DDC_State += "KO"
														    $DC += ,$item
													    }
												    }
												    else
												    {
													    $DDC_State += "KO"
													    $DC += ,$item
												    }
											    }
											    $Test_DB = (Test-BrokerDBConnection (Get-BrokerDBConnection -AdminAddress $SyncHash_Conf.DDC_TB)).ServiceStatus
											    $Farm_infos = New-Object -TypeName PSObject -Property @{ "Farm" = $Farm; "Version" = $Version; "DDC" = $DC; "DDC_State" = $DDC_State; "Farm_State" = $Test_DB }
											    $Farm_List_Conf = New-Object System.Collections.Generic.List[Object]
											    $Farm_List_Conf.Add($Farm_Infos)
											    if ($SyncHash_Conf.ListView_Conf.itemssource.farm -contains $Farm)
											    {
												    $SyncHash_Conf.Form_Configuration.Dispatcher.Invoke([action]{
														    $SyncHash_Conf.ApplicationLayer.IsEnabled = $false
														    $SyncHash_Conf.SpinnerOverlayLayer.Visibility = "Collapsed"
														    $SyncHash_Conf.DDC_MB.Foreground = "Red"
														    $SyncHash_Conf.DDC_MB.FontSize = "20"
														    $SyncHash_Conf.DDC_MB.text = "Farm $Farm aleady added.`r`nPlease enter a DDC from a new farm."
														    $SyncHash_Conf.Dialog_DDC.IsOpen = $True
														    $SyncHash_Conf.DDC_State.source = $SyncHash_Conf.Grey
													    }, "Normal")
											    }
											    Else
											    {
												    $SyncHash_Conf.ListView_Conf.Dispatcher.Invoke([Action]{ $SyncHash_Conf.ListView_Conf.ItemsSource += $Farm_List_Conf }, [Windows.Threading.DispatcherPriority]::Normal)
												    $SyncHash_Conf.Form_Configuration.Dispatcher.Invoke([action]{
														    $SyncHash_Conf.ApplicationLayer.IsEnabled = $false
														    $SyncHash_Conf.SpinnerOverlayLayer.Visibility = "Collapsed"
														    $SyncHash_Conf.DDC_MB.Foreground = "Blue"
														    $SyncHash_Conf.DDC_MB.FontSize = "20"
														    $SyncHash_Conf.DDC_MB.text = "Farm $Farm added."
														    $SyncHash_Conf.Dialog_DDC.IsOpen = $True
														    $SyncHash_Conf.DDC_State.source = $SyncHash_Conf.Grey
													    }, "Normal")
											    }
										    }
									    }
									    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Add_DDC_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
								    })
							    Worker
							    $DDC_MB_Close.add_Click({
									    $Dialog_DDC.IsOpen = $False
									    $ApplicationLayer.IsEnabled = $true
								    })
						    }
						    $FarmColumnIndex = -1
						    for ($i = 0; $i -lt $ListView_Conf.View.Columns.Count; $i++)
						    {
							    if ($ListView_Conf.View.Columns[$i].Header -eq "Farm")
							    {
								    $FarmColumnIndex = $i
								    break
							    }
						    }
						    if ($FarmColumnIndex -ne -1)
						    {
							    $ListView_Conf.Items.SortDescriptions.Clear()
							    $ListView_Conf.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription "Farm", "Ascending"))
							    $ListView_Conf.Items.Refresh()
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Add_DDC " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    ### Configuration // Remove_DDC
			    $Remove_DDC.add_Click({
					    try
					    {
						    if ($ListView_Conf.SelectedItem -eq $null)
						    {
							    $ApplicationLayer.IsEnabled = $false
							    $DDC_MB.Foreground = "Red"
							    $DDC_MB.FontSize = "20"
							    $DDC_MB.text = "Please select a farm."
							    $Dialog_DDC.IsOpen = $True
							    $DDC_MB_Close.add_Click({
									    $Dialog_DDC.IsOpen = $False
									    $ApplicationLayer.IsEnabled = $true
								    })
						    }
						    Else
						    {
							    $Farm_to_R = $ListView_Conf.SelectedItem.Farm
							    $Farm_List_Conf_R = [System.Collections.ObjectModel.ObservableCollection[Object]]$ListView_Conf.ItemsSource
							    $Farm_List_Conf_R.Remove($ListView_Conf.SelectedItem)
							    $ListView_Conf.ItemsSource = $Farm_List_Conf_R
							    $DDC_MB.Foreground = "Red"
							    $DDC_MB.FontSize = "20"
							    $DDC_MB.text = "Farm $Farm_to_R removed."
							    $Dialog_DDC.IsOpen = $True
							    $DDC_MB_Close.add_Click({
									    $Dialog_DDC.IsOpen = $False
									    $ApplicationLayer.IsEnabled = $true
								    })
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Remove_DDC " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    ### Configuration // Finish_DDC
			    $Finish_DDC.add_Click({
					    try
					    {
						    if ($ListView_Conf.ItemsSource.count -eq 0)
						    {
							    $ApplicationLayer.IsEnabled = $false
							    $DDC_MB.Foreground = "Red"
							    $DDC_MB.FontSize = "20"
							    $DDC_MB.text = "No farm added.`r`nPlease add at least one farm`r`nor cancel configuration."
							    $Dialog_DDC.IsOpen = $True
							    $DDC_MB_Close.add_Click({
									    $Dialog_DDC.IsOpen = $False
									    $ApplicationLayer.IsEnabled = $true
								    })
						    }
						    Else
						    {
							    $Form_Configuration.Close()
							    $Form_Check_conf.Hide()
							    if (Test-Path $global:LogoFile) { Remove-Item -Path $global:LogoFile -Force }
							    if (Test-Path $ConfigFile) { Remove-Item -Path $ConfigFile -Force }
							    if (!(Test-Path $ConfigPath)) { New-Item -Path $ConfigPath -Type Directory }
							    if (Test-Path $global:SelectLogoFile) { Copy-Item $global:SelectLogoFile $global:LogoFile -Force }
							    $ListView_Conf.ItemsSource | Export-Clixml -Path $ConfigFile
							    $datas = Import-Clixml -Path $ConfigFile
							    $Load_TB.Text = "Loading configuration file"
							    if (Test-Path variable:\$SyncHash.Farm_List) { $SyncHash.Farm_List = @() }
							    $S_Sessions.Items.Clear()
							    $S_VDAs.Items.Clear()
							    $S_Publications.Items.Clear()
							    $S_Sessions_Details.Items.Clear()
							    $S_Publications_Details.Items.Clear()
							    $S_VDAs_Details.Items.Clear()
							    $S_MCs_Details.Items.Clear()
							    $S_DGs_Details.Items.Clear()
							    $S_Maintenance.Items.Clear()
							    $S_Registration.Items.Clear()
							    $Tab_Control.SelectedIndex = 0
							    Process-FarmData -datas $datas
							    $Form.Add_Closing({
									    $Process = Get-Process XD_Tool -ErrorAction SilentlyContinue
									    if ($Process) { $Process | Stop-Process -Force }
								    })
							    $Form.ShowDialog() | Out-Null
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Finish_DDC " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    ### Configuration // Cancel_DDC
			    $Cancel_DDC.add_Click({
					    try
					    {
						    $Farm_Infos = @()
						    $Farm_List_Conf = @()
						    $DDC_TB.text = ""
						    $Logo_Conf.Source = $null
						    if ($ListView_Conf.ItemsSource.count -ne 0)
						    {
							    $ListView_Conf.ItemsSource.Clear()
							    $ListView_Conf.ItemsSource = $null
							    $ListView_Conf.ItemsSource = @()
						    }
						    $Form_Configuration.Close()
						    $Form_Check_conf.ShowDialog()
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Cancel_DDC " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    ### Configuration // Logo
			    $Add_Logo_Conf.add_Click({
					    try
					    {
						    $fileDialog = New-Object Microsoft.Win32.OpenFileDialog
						    $fileDialog.Filter = "Image files (*.jpg, *.jpeg, *.png, *.gif)|*.jpg;*.jpeg;*.png;*.gif"
						    $fileDialog.Multiselect = $false
						    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
						    $result = $fileDialog.ShowDialog()
						    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
						    {
							    $global:SelectLogoFile = $fileDialog.FileName
							    $source = New-Object System.Windows.Media.Imaging.BitmapImage
							    $source.BeginInit()
							    $source.UriSource = New-Object System.Uri($fileDialog.FileName)
							    $source.EndInit()
							    $Logo_Conf.Source = $source
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Add_Logo_Conf " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    ############################
			    ### End_Configuration // New
			    ############################
			    $Form_Check_conf.Hide()
			    $DDC_State.source = $Grey
			    $Form_Configuration.ShowDialog()
			
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_New_Conf " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    ##################################
    ### End_Check_Configuration // New
    ##################################
    ### Check_Configuration // Cancel
    $Cancel_Conf.add_Click({
		    try
		    {
			    $Form_Check_conf.Hide()
			    if ($Form.Visibility -eq "Hidden")
			    {
				    $Form.Add_Closing({
						    $Process = Get-Process XD_Tool -ErrorAction SilentlyContinue
						    if ($Process) { $Process | Stop-Process -Force }
					    })
				    $Form.ShowDialog()
			    }
			    else { Exit }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Cancel_Conf " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    ###########################
    ### End_Check_Configuration
    ###########################
    ############
    # Start_MAIN
    ############
    $Refresh_Main.add_Click({
		    try
		    {
			    $datas = Import-Clixml -Path $ConfigFile
			    $Load_TB.Text = "Refresh configuration data"
			    Process-FarmData -datas $datas
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Main " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Licenses_Help.add_Click({ $Snackbar.MessageQueue.Enqueue("Green : left licenses greater than 5%. - Orange : left licenses between 3% and 5%. - Red : left licenses less than 3%.") })
    $S_License.Add_SelectionChanged({
		    try
		    {
			    $selectedItem = $S_License.SelectedItem
			    $LicVar = $SyncHash.Keys | ? {$SyncHash[$_].ProductName -eq $selectedItem}
                $VarObj = $SyncHash.$LicVar
				$Percent_Lic = [math]::Round((($VarObj.Left/$VarObj.Count) * 100), 2)
				if ($Percent_Lic -lt 3) { $leftColor = "Red" }
				elseif ($Percent_Lic -gt 3 -and $Percent_Lic -lt 5) { $leftColor = "Orange" }
				else { $leftColor = "Green" }
				$Licenses_TB.Text = "License Edition : $($VarObj['LicenseEdition'])`nLicense Model : $($VarObj['LicenseModel'])`nLocalized License Model : $($VarObj['LocalizedLicenseModel'])`nSubscription Advantage Date : $($VarObj['SubscriptionAdvantageDate'])`nLicense Expiration Date : $($VarObj['LicenseExpirationDate'])`n$($VarObj['Count']) purchased`n$($VarObj['InUseCount']) used`n"
				$Licenses_TB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = $VarObj.Left; Foreground = $LeftColor }))
				$Licenses_TB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = " left" }))
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_License " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Licenses_Refresh.add_Click({
		    try
		    {
			    $SyncHash.S_License.Items.Clear()
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing licenses status"
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $DDC_License = $SyncHash.($SyncHash.Farm_List[0]).DDC
						    $License_server = (Get-BrokerSite -AdminAddress $DDC_License).LicenseServerName
					        $CertHash = (Get-LicCertificate -AdminAddress $License_server).CertHash
					        $Lic_Inventory = Get-LicInventory -AdminAddress $License_server -CertHash $CertHash | Sort-Object -Descending LicenseProductName
						    $Lic_List = @()
						    $SyncHash.Lic_Var = @()
						    foreach ($Type in $Lic_Inventory)
						    {
							    $LicenseProductName = $Type.LicenseProductName
							    $VariableName = "${LicenseProductName}"
							    $LocalizedLicenseProductName = $Type.LocalizedLicenseProductName
							    $LicenseEdition = $Type.LicenseEdition
							    $LicenseSubscriptionAdvantageDate = ($Type.LicenseSubscriptionAdvantageDate).ToString('yyyy.MMdd')
                                $LicenseExpirationDate = $Type.LicenseExpirationDate
                                $LocalizedLicenseModel = $Type.LocalizedLicenseModel
							    [int]$InUseCount = $Type.LicensesInUse
							    [int]$Count = $Type.LicensesAvailable
							    [int]$Left = $Count - $InUseCount
							    $LicenseModel = $Type.LicenseModel
							    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
							    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
							    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
							    $VariableValue = @{ ProductName = $LocalizedLicenseProductName; LicenseEdition = if ($LicenseEdition.Length -eq 0) { "N/A" } else {$LicenseEdition}; SubscriptionAdvantageDate = $LicenseSubscriptionAdvantageDate; LicenseExpirationDate = $LicenseExpirationDate; LocalizedLicenseModel = $LocalizedLicenseModel; InUseCount = $InUseCount; Count = $Count; LicenseModel = if ($LicenseModel.Length -eq 0) { "N/A" } else {$LicenseModel}; Left = $Left }
							    New-Variable -Name $VariableName -Value $VariableValue
							    $SyncHash.Add($VariableName, $VariableValue)
							    $Lic_List += $Type.LocalizedLicenseProductName
							    $SyncHash.Lic_Var += $VariableName
						    }
						    foreach ($item in $Lic_List) { $SyncHash.S_License.Dispatcher.Invoke([Action]{ $SyncHash.S_License.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $true
								    $SyncHash.S_License.SelectedItem = $Lic_List[0]
							    }, "Normal")
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Licenses_Refresh_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Licenses_Refresh " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Sessions_Help.add_Click({ $Snackbar.MessageQueue.Enqueue("Session state 'connected' is an issue.") })
    $S_Sessions.Add_SelectionChanged({
		    try
		    {
			    $selectedItem = $S_Sessions.SelectedItem
			    if ($SyncHash.S_Sessions.Items.count -match 0) { $Sessions_TB.Text = "Total sessions : `nActive Sessions : `nDisconnected Sessions : `nConnected Sessions : " }
			    else
			    {
				    $select = "Sessions_" + $selectedItem
				    $VarObj = $SyncHash.$select
				    if ($($VarObj['Connected_Sessions']) -eq 0) { $Connected_Sessions_Color = "Green" }
				    else { $Connected_Sessions_Color = "Red" }
				    $Sessions_TB.Text = "Total sessions : $($VarObj['Total_Sessions'])`nActive Sessions : $($VarObj['Active_Sessions'])`nDisconnected Sessions : $($VarObj['Disconnected_Sessions'])`nConnected Sessions : "
				    $Sessions_TB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = $VarObj.Connected_Sessions; Foreground = $Connected_Sessions_Color }))
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_Sessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Sessions_Refresh.add_Click({
		    try
		    {
			    $SyncHash.S_Sessions.Items.Clear()
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing sessions status"
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $SyncHash.Total_Sessions_All = $null
						    $SyncHash.Active_Sessions_All = $null
						    $SyncHash.Disconnected_Sessions_All = $null
						    $SyncHash.Connected_Sessions_All = $null
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($Farm in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$Farm).DDC
							    $Sessions = Get-BrokerSession -MaxRecordCount 999999 -AdminAddress $DDC
							    $Total_Sessions = $Sessions.count
							    $SyncHash.Total_Sessions_All += $Sessions.count
							    $Active_Sessions = ($Sessions | Where-Object { $_.SessionState -eq "Active" }).count
							    $SyncHash.Active_Sessions_All += ($Sessions | Where-Object { $_.SessionState -eq "Active" }).count
							    $Disconnected_Sessions = ($Sessions | Where-Object { $_.SessionState -eq "Disconnected" }).count
							    $SyncHash.Disconnected_Sessions_All += ($Sessions | Where-Object { $_.SessionState -eq "Disconnected" }).count
							    $Connected_Sessions = ($Sessions | Where-Object { $_.SessionState -eq "Connected" }).count
							    $SyncHash.Connected_Sessions_All += ($Sessions | Where-Object { $_.SessionState -eq "Connected" }).count
							    $VariableName = "Sessions_" + "${Farm}"
							    $VariableValue = @{ Total_Sessions = $Total_Sessions; Active_Sessions = $Active_Sessions; Disconnected_Sessions = $Disconnected_Sessions; Connected_Sessions = $Connected_Sessions }
							    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
							    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
							    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
							    New-Variable -Name $VariableName -Value $VariableValue
							    $SyncHash.Add($VariableName, $VariableValue)
						    }
						    $VariableName = "Sessions_All farms"
						    $VariableValue = @{ Total_Sessions = $SyncHash.Total_Sessions_All; Active_Sessions = $SyncHash.Active_Sessions_All; Disconnected_Sessions = $SyncHash.Disconnected_Sessions_All; Connected_Sessions = $SyncHash.Connected_Sessions_All }
						    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
						    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
						    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
						    New-Variable -Name $VariableName -Value $VariableValue
						    $SyncHash.Add($VariableName, $VariableValue)
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Sort-Object
						    if ($SyncHash.Farm_List.count -ne 1 -and $SyncHash.Farm_List -notcontains "All farms")
						    {
							    $Array = @("All farms") + $SyncHash.Farm_List
							    $SyncHash.Farm_List = $Array
						    }
						    else { $SyncHash.Farm_List = ,$SyncHash.Farm_List }
						    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.S_Sessions.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $true
								    $SyncHash.S_Sessions.SelectedItem = $SyncHash.S_Sessions.Items[0]
							    }, "Normal")
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Sessions_Refresh_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Sessions_Refresh " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $S_VDAs.Add_SelectionChanged({
		    try
		    {
			    $selectedItem = $S_VDAs.SelectedItem
			    if ($SyncHash.S_VDAs.Items.count -match 0) { $VDAs_TB.Text = "Total VDAs : `nServers : `nVDIs : `nPowered Off : `nMaintenance : `nUnregistered : " }
			    else
			    {
				    $select = "VDAs_" + $selectedItem
				    $VarObj = $SyncHash.$select
				    if ($($VarObj['Unregistered']) -eq 0) { $Unregistered_Color = "Green" }
				    else { $Unregistered_Color = "Red" }
				    $VDAs_TB.Text = "Total VDAs : $($VarObj['Total_VDAs'])`nServers : $($VarObj['Total_Servers'])`nVDIs : $($VarObj['Total_VDIs'])`nPoweredOff : $($VarObj['PoweredOff'])`nMaintenance : $($VarObj['Maintenance'])`nUnregistered : "
				    $VDAs_TB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = $VarObj.Unregistered; Foreground = $Unregistered_Color }))
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_VDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $VDAs_Refresh.add_Click({
		    try
		    {
			    $SyncHash.S_VDAs.Items.Clear()
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing VDAs status"
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $SyncHash.Total_VDAs_All = $null
						    $SyncHash.Total_Servers_All = $null
						    $SyncHash.Total_VDIs_All = $null
						    $SyncHash.PoweredOff_All = $null
						    $SyncHash.Maintenance_All = $null
						    $SyncHash.Unregistered_All = $null
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($Farm in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$Farm).DDC
							    $VDAs = Get-BrokerDesktop -MaxRecordCount 999999 -AdminAddress $DDC
							    $Total_VDAs = $VDAs.count
							    $SyncHash.Total_VDAs_All += $VDAs.count
							    $Total_Servers = ($VDAs | Where-Object { $_.DesktopKind -eq "Shared" }).count
							    $SyncHash.Total_Servers_All += ($VDAs | Where-Object { $_.DesktopKind -eq "Shared" }).count
							    $Total_VDIs = ($VDAs | Where-Object { $_.DesktopKind -eq "Private" }).count
							    $SyncHash.Total_VDIs_All += ($VDAs | Where-Object { $_.DesktopKind -eq "Private" }).count
							    $PoweredOff = ($VDAs | Where-Object { $_.PowerState -eq "Off" }).count
							    $SyncHash.PoweredOff_All += ($VDAs | Where-Object { $_.PowerState -eq "Off" }).count
							    $Maintenance = ($VDAs | Where-Object { $_.InMaintenanceMode -eq $true }).count
							    $SyncHash.Maintenance_All += ($VDAs | Where-Object { $_.InMaintenanceMode -eq $true }).count
							    $Unregistered = ($VDAs | Where-Object { $_.RegistrationState -eq "Unregistered" }).count
							    $SyncHash.Unregistered_All += ($VDAs | Where-Object { $_.RegistrationState -eq "Unregistered" }).count
							    $VariableName = "VDAs_" + "${Farm}"
							    $VariableValue = @{ Total_VDAs = $Total_VDAs; Total_Servers = $Total_Servers; Total_VDIs = $Total_VDIs; PoweredOff = $PoweredOff; Maintenance = $Maintenance; Unregistered = $Unregistered }
							    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
							    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
							    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
							    New-Variable -Name $VariableName -Value $VariableValue
							    $SyncHash.Add($VariableName, $VariableValue)
						    }
						    $VariableName = "VDAs_All farms"
						    $VariableValue = @{ Total_VDAs = $SyncHash.Total_VDAs_All; Total_Servers = $SyncHash.Total_Servers_All; Total_VDIs = $SyncHash.Total_VDIs_All; PoweredOff = $SyncHash.PoweredOff_All; Maintenance = $SyncHash.Maintenance_All; Unregistered = $SyncHash.Unregistered_All }
						    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
						    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
						    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
						    New-Variable -Name $VariableName -Value $VariableValue
						    $SyncHash.Add($VariableName, $VariableValue)
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Sort-Object
						    if ($SyncHash.Farm_List.count -ne 1 -and $SyncHash.Farm_List -notcontains "All farms")
						    {
							    $Array = @("All farms") + $SyncHash.Farm_List
							    $SyncHash.Farm_List = $Array
						    }
						    else { $SyncHash.Farm_List = ,$SyncHash.Farm_List }
						    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_VDAs.Dispatcher.Invoke([Action]{ $SyncHash.S_VDAs.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $true
								    $SyncHash.S_VDAs.SelectedItem = $SyncHash.S_VDAs.Items[0]
							    }, "Normal")
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDAs_Refresh_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDAs_Refresh " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $S_Publications.Add_SelectionChanged({
		    try
		    {
			    $selectedItem = $S_Publications.SelectedItem
			    if ($SyncHash.S_Publications.Items.count -match 0) { $Publications_TB.Text = "Total Publications : `nTotal Applications : `nTotal Desktops : `nPublications Disabled : `nPublications Hidden : " }
			    else
			    {
				    $select = "Publications_" + $selectedItem
				    $VarObj = $SyncHash.$select
				    $Publications_TB.Text = "Total Publications : $($VarObj['Total_Publications'])`nTotal Applications : $($VarObj['Total_Applications'])`nTotal Desktops : $($VarObj['Total_Desktops'])`nPublications Disabled : $($VarObj['Publications_Disabled'])`nPublications Hidden : $($VarObj['Publications_Hidden'])"
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_Publications " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Publications_Refresh.add_Click({
		    try
		    {
			    $SyncHash.S_Publications.Items.Clear()
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing Publications status"
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $SyncHash.Total_Publications_All = $null
						    $SyncHash.Total_Applications_All = $null
						    $SyncHash.Total_Desktops_All = $null
						    $SyncHash.Publications_Disabled_All = $null
						    $SyncHash.Publications_Hidden_All = $null
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
						    foreach ($Farm in $SyncHash.Farm_List)
						    {
							    $DDC = ($SyncHash.$Farm).DDC
							    $Applications = Get-BrokerApplication -MaxRecordCount 999999 -AdminAddress $DDC
							    $Desktops = Get-BrokerEntitlementPolicyRule -MaxRecordCount 999999 -AdminAddress $DDC
							    $Total_Applications = $Applications.count
							    $SyncHash.Total_Applications_All += $Total_Applications
							    $Total_Desktops = $Desktops.count
							    $SyncHash.Total_Desktops_All += $Total_Desktops
							    $Total_Publications = $Total_Applications + $Total_Desktops
							    $SyncHash.Total_Publications_All += $Total_Publications
							    $Publications_Disabled = ($Applications | Where-Object { $_.Enabled -eq $False }).count
							    $SyncHash.Publications_Disabled_All += $Publications_Disabled
							    $Publications_Hidden = ($Applications | Where-Object { $_.Visible -eq $False }).count
							    $SyncHash.Publications_Hidden_All += $Publications_Hidden
							    $VariableName = "Publications_" + "${Farm}"
							    $VariableValue = @{ Total_Publications = $Total_Publications; Total_Applications = $Total_Applications; Total_Desktops = $Total_Desktops; Publications_Disabled = $Publications_Disabled; Publications_Hidden = $Publications_Hidden }
							    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
							    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
							    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
							    New-Variable -Name $VariableName -Value $VariableValue
							    $SyncHash.Add($VariableName, $VariableValue)
						    }
						    $VariableName = "Publications_All farms"
						    $VariableValue = @{ Total_Publications = $SyncHash.Total_Publications_All; Total_Applications = $SyncHash.Total_Applications_All; Total_Desktops = $SyncHash.Total_Desktops_All; Publications_Disabled = $SyncHash.Publications_Disabled_All; Publications_Hidden = $SyncHash.Publications_Hidden_All }
						    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
						    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
						    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
						    New-Variable -Name $VariableName -Value $VariableValue
						    $SyncHash.Add($VariableName, $VariableValue)
						    $SyncHash.Farm_List = $SyncHash.Farm_List | Sort-Object
						    if ($SyncHash.Farm_List.count -ne 1 -and $SyncHash.Farm_List -notcontains "All farms")
						    {
							    $Array = @("All farms") + $SyncHash.Farm_List
							    $SyncHash.Farm_List = $Array
						    }
						    else { $SyncHash.Farm_List = ,$SyncHash.Farm_List }
						    foreach ($item in $SyncHash.Farm_List) { $SyncHash.S_Publications.Dispatcher.Invoke([Action]{ $SyncHash.S_Publications.Items.Add($item) }, [Windows.Threading.DispatcherPriority]::Normal) }
						
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $true
								    $SyncHash.S_Publications.SelectedItem = $SyncHash.S_Publications.Items[0]
							    }, "Normal")
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publications_Refresh_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publications_Refresh " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Farm_State_Help.add_Click({ $snackbar.MessageQueue.Enqueue("The Farm State is OK if at least one DDC is up and able to join the database.") })
    $Farms_Refresh.add_Click({
		    try
		    {
			    $SyncHash.ListView_Main.ItemsSource.Clear()
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing farms status"
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $Farm_Infos = @()
						    $SyncHash.Farm_List = @()
						    $Farm_ListView = New-Object System.Collections.Generic.List[Object]
						    foreach ($data in $SyncHash.datas)
						    {
							    $DC_State = @()
							    $FarmName = $data.Farm
							    $Version = $data.Version
							    if (Test-Path Variable:\$FarmName) { Remove-Variable $FarmName }
							    $DDC = @()
							    ForEach ($DC in $data.DDC -split "`n")
							    {
								    if ((Test-Connection $DC -Count 1 -ErrorAction SilentlyContinue) -and $DC -match $env:COMPUTERNAME)
								    {
									    if ((Get-Service -Name "CitrixBrokerService").Status -eq "Running")
									    {
										    $DC_State += "OK"
										    $DDC += ,$DC
									    }
									    else
									    {
										    $DC_State += "KO"
										    $DDC += ,$DC
									    }
								    }
								    elseif ((Test-Connection $DC -Count 1 -ErrorAction SilentlyContinue) -and $DC -notmatch $env:COMPUTERNAME)
								    {
									    if (((Invoke-Command -ComputerName $DC -ScriptBlock { Get-Service -Name "CitrixBrokerService" }).Status) -eq "Running")
									    {
										    $DC_State += "OK"
										    $DDC += ,$DC
									    }
									    else
									    {
										    $DC_State += "KO"
										    $DDC += ,$DC
									    }
								    }
								    else
								    {
									    $DC_State += "KO"
									    $DDC += ,$DC
								    }
							    }
							    ForEach ($DC in $data.DDC -split "`n")
							    {
								    if ((Test-Connection $DC -Count 1 -ErrorAction SilentlyContinue) -and $DC -match $env:COMPUTERNAME)
								    {
									    if ((Get-Service -Name "CitrixBrokerService").Status -eq "Running")
									    {
										    $Test_DB = (Test-BrokerDBConnection (Get-BrokerDBConnection -AdminAddress $DC)).ServiceStatus
										    if ($Test_DB -eq "OK")
										    {
											    $VariableName = "${FarmName}"
											    $VariableValue = @{ Farm = $FarmName; DDC = $DC }
											    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
											    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
											    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
											    New-Variable -Name $VariableName -Value $VariableValue
											    $SyncHash.Add($VariableName, $VariableValue)
											    $SyncHash.Farm_List = $SyncHash.Farm_List + $VariableName
											    break
										    }
									    }
									    else { $Test_DB = "KO" }
								    }
								    elseif ((Test-Connection $DC -Count 1 -ErrorAction SilentlyContinue) -and $DC -notmatch $env:COMPUTERNAME)
								    {
									    if (((Invoke-Command -ComputerName $DC -ScriptBlock { Get-Service -Name "CitrixBrokerService" }).Status) -eq "Running")
									    {
										    $Test_DB = (Test-BrokerDBConnection (Get-BrokerDBConnection -AdminAddress $DC)).ServiceStatus
										    if ($Test_DB -eq "OK")
										    {
											    $VariableName = "${FarmName}"
											    $VariableValue = @{ Farm = $FarmName; DDC = $DC }
											    if (Test-Path Variable:\$VariableName) { Remove-Variable $VariableName }
											    if ($SyncHash.ContainsKey($VariableName)) { $SyncHash.Remove($VariableName) }
											    if (Test-Path Variable:\SyncHash) { $SyncHash.VariableName = $null }
											    New-Variable -Name $VariableName -Value $VariableValue
											    $SyncHash.Add($VariableName, $VariableValue)
											    $SyncHash.Farm_List = $SyncHash.Farm_List + $VariableName
											    break
										    }
									    }
									    else { $Test_DB = "KO" }
								    }
								    else { $Test_DB = "KO" }
							    }
							    $Farm_infos = New-Object -TypeName PSObject -Property @{ "Farm" = $data.Farm; "Version" = $data.Version; "DDC" = $DDC; "DDC_State" = $DC_State; "Farm_State" = $Test_DB }
							    $Farm_ListView.Add($Farm_Infos)
						    }
						    $SyncHash.Form.Dispatcher.Invoke([action]{
								    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
								    $SyncHash.MainLayer.IsEnabled = $true
								    $SyncHash.ListView_Main.ItemsSource = $Farm_ListView
							    }, "Normal")
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Farms_Refresh_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $FarmColumnIndex = -1
			    for ($i = 0; $i -lt $ListView_Main.View.Columns.Count; $i++)
			    {
				    if ($ListView_Main.View.Columns[$i].Header -eq "Farm")
				    {
					    $FarmColumnIndex = $i
					    break
				    }
			    }
			    if ($FarmColumnIndex -ne -1)
			    {
				    $ListView_Main.Items.SortDescriptions.Clear()
				    $ListView_Main.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription "Farm", "Ascending"))
				    $ListView_Main.Items.Refresh()
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Farms_Refresh " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    ##########
    # End_MAIN
    ##########
    ################
    # Start_Sessions
    ################
    $UserName_TB.Add_KeyDown({
		    param ($sender,
			    $e)
		    if ($e.Key -eq [System.Windows.Input.Key]::Enter) { Search_User }
	    })
    $Search_User_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for a user sessions with username, first name or last name. ") })
    $Search_Sessions_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for all sessions in the farm selected.") })
    $Search_User.add_Click({ Search_User })
    $Search_sessions.add_Click({
		    try
		    {
			    if ($datagrid_usersList.itemsSource.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please search a user." }
			    elseif ($datagrid_usersList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a user." }
			    else
			    {
				    $Load_TB.Text = "Searching sessions"
				    Refresh_Session
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Search_sessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Kill_Session.add_Click({
		    try
		    {
			    $Sessions = $datagrid_UserSessions.SelectedItems
			    if ($datagrid_usersList.itemsSource.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please search a user." }
			    elseif ($datagrid_UserSessions.ItemsSource -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please search for a session." }
			    elseif ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    Else
			    {
				    foreach ($Session in $Sessions)
				    {
					    $Farm = $Session.Farm
					    $UID = $Session.Uid
					    $DDC = ($SyncHash.$Farm).DDC
					    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Stop-BrokerSession
				    }
				    Show-Dialog_Main -Foreground "Blue" -Text "Kill command sent.`r`nPlease refresh sessions in few seconds."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Kill_Session " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Hide_Session.add_Click({
		    try
		    {
			    $Sessions = $datagrid_UserSessions.SelectedItems
			    if ($datagrid_usersList.itemsSource.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please search a user." }
			    elseif ($datagrid_UserSessions.ItemsSource -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please search for a session." }
			    elseif ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    else
			    {
				    $VDI = 0
				    $i = 0
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = $SyncHash.($Session.Farm).DDC
					    $Hidden = $Session.Hidden
					    $Type = $Session.Type
					    if ($Type -match "VDI") { $VDI += 1 }
					    else
					    {
						    if ($Hidden -eq $false)
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$true
						    }
						    else
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$false
						    }
					    }
				    }
				    if ($VDI -ne 0 -and $i -ne 0)
				    {
					    $Load_TB.Text = "Refreshing sessions"
					    Refresh_Session
					    $MainLayer.IsEnabled = $false
					    $Main_MB.FontSize = "20"
					    $Main_MB.text = "Hide status changed for server connections."
					    $Main_MB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = "`r`nBut hide status for VDI can't be changed."; Foreground = "Red" }))
					    $Dialog_Main.IsOpen = $True
					    $Main_MB_Close.add_Click({
							    $Dialog_Main.IsOpen = $False
							    $MainLayer.IsEnabled = $true
						    })
				    }
				    if ($VDI -ne 0 -and $i -eq 0) { Show-Dialog_Main -Foreground "Red" -Text "Hide status for VDI can't be changed." }
				    if ($VDI -eq 0 -and $i -ne 0)
				    {
					    $Load_TB.Text = "Refreshing sessions"
					    Refresh_Session
					    Show-Dialog_Main -Foreground "Blue" -Text "Hide status changed."
				    }
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Hide_Session " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Shadow_Session.add_Click({
		    try
		    {
			    $Session = $datagrid_UserSessions.SelectedItems
			    if ($datagrid_usersList.itemsSource.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please search a user." }
			    elseif ($datagrid_UserSessions.ItemsSource -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please search for a session." }
			    elseif ($Session.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    elseif ($Session.count -ge "2") { Show-Dialog_Main -Foreground "Red" -Text "Please select only one session." }
			    else
			    {
				    $Machine = $Session."Machine Name"
				    $User = $Session.User
				    $Domain = $Session.Domain
				    $ID = Get-UserNameSessionIDMap -Comp $Machine | ? { $_.UserName -match $User } | Select-Object -ExpandProperty SessionID
				    $Arg = "/offerra $Machine "
				    $Arg += "$Domain\$User"
				    $Arg += ":"
				    $Arg += $ID
				    Start-Process msra $Arg
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Shadow_Session " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_Session.add_Click({
		    try
		    {
			    if ($datagrid_UserSessions.itemsSource.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Nothing to refresh." }
			    else
			    {
				    $Load_TB.Text = "Refreshing sessions"
				    Refresh_Session
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Session " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $S_Sessions_Details.Add_SelectionChanged({
		    try
		    {
			    if ($S_Sessions_Details.selectedItem -ne $null)
			    {
				    $UserName_TB.text = ""
				    $Load_TB.Text = "Searching sessions"
				    $datagrid_usersList.Visibility = "Collapsed"
				    $Search_sessions.Visibility = "Collapsed"
				    $datagrid_UserSessions.Visibility = "Collapsed"
				    $Kill_Session.Visibility = "Collapsed"
				    $Hide_Session.Visibility = "Collapsed"
				    $Shadow_Session.Visibility = "Collapsed"
				    $Refresh_Session.Visibility = "Collapsed"
				    Refresh_AllSessions
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_Sessions_Details " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Kill_AllSessions.add_Click({
		    try
		    {
			    $Sessions = $datagrid_AllSessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    Else
			    {
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = ($SyncHash.$Session.Farm).DDC
					    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Stop-BrokerSession
				    }
				    Show-Dialog_Main -Foreground "Blue" -Text "Kill command sent.`r`nPlease refresh sessions in few seconds."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Kill_AllSessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Hide_AllSessions.add_Click({
		    try
		    {
			    $Sessions = $datagrid_AllSessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    else
			    {
				    $VDI = 0
				    $i = 0
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = $SyncHash.($Session.Farm).DDC
					    $Hidden = $Session.Hidden
					    $Type = $Session.Type
					    if ($Type -match "VDI") { $VDI += 1 }
					    else
					    {
						    if ($Hidden -eq $false)
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$true
						    }
						    else
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$false
						    }
					    }
				    }
				    if ($VDI -ne 0 -and $i -ne 0)
				    {
					    $Load_TB.Text = "Refreshing sessions"
					    Refresh_AllSessions
					    $MainLayer.IsEnabled = $false
					    $Main_MB.FontSize = "20"
					    $Main_MB.text = "Hide status changed for server connections."
					    $Main_MB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = "`r`nBut hide status for VDI can't be changed."; Foreground = "Red" }))
					    $Dialog_Main.IsOpen = $True
					    $Main_MB_Close.add_Click({
							    $Dialog_Main.IsOpen = $False
							    $MainLayer.IsEnabled = $true
						    })
				    }
				    if ($VDI -ne 0 -and $i -eq 0) { Show-Dialog_Main -Foreground "Red" -Text "Hide status for VDI can't be changed." }
				    if ($VDI -eq 0 -and $i -ne 0)
				    {
					    $Load_TB.Text = "Refreshing sessions"
					    Refresh_AllSessions
					    Show-Dialog_Main -Foreground "Blue" -Text "Hide status changed."
				    }
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Hide_AllSessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Shadow_AllSessions.add_Click({
		    try
		    {
			    $Session = $datagrid_AllSessions.SelectedItems
			    if ($Session.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    elseif ($Session.count -ge "2") { Show-Dialog_Main -Foreground "Red" -Text "Please select only one session." }
			    else
			    {
				    $Machine = $Session."Machine Name"
				    $User = $Session.User
				    $Domain = $Session.Domain
				    $ID = Get-UserNameSessionIDMap -Comp $Machine | ? { $_.UserName -match $User } | Select-Object -ExpandProperty SessionID
				    $Arg = "/offerra $Machine "
				    $Arg += "$Domain\$User"
				    $Arg += ":"
				    $Arg += $ID
				    Start-Process msra $Arg
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Shadow_AllSessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_AllSessions.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing sessions"
			    Refresh_AllSessions
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllSessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_AllSessions.add_Click({
		    try
		    {
			    if ($datagrid_AllSessions.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No sessions to export." }
			    else
			    {
				    $FarmSelected = $S_Sessions_Details.selectedItem
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportSessions = [hashtable]::Synchronized(@{
						    FarmSelected = $FarmSelected
						    ConfigPath   = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportSessions", $SyncHash_ExportSessions)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $Farm = $SyncHash_ExportSessions.FarmSelected
							    $date = get-date -Format MM_dd_yyyy
							    $Export_Sessions = $SyncHash_ExportSessions.ConfigPath + "\Exports\Sessions_$Farm-$date.xlsx"
							    while (Test-Path $Export_Sessions)
							    {
								    $i++
								    $Export_Sessions = $SyncHash_ExportSessions.ConfigPath + "\Exports\Sessions_$Farm-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_AllSessions.ItemsSource | Export-xlsx -Path $Export_Sessions -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_Sessions"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_AllSessions_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_AllSessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_AllSessions_Simple.add_Click({
		    $Grid_AllSessions_Simple.Visibility = "visible"
		    $Grid_AllPSessions_Full.Visibility = "collapse"
	    })
    $Switch_AllSessions_Full.add_Click({
		    $Grid_AllSessions_Simple.Visibility = "collapse"
		    $Grid_AllPSessions_Full.Visibility = "visible"
	    })
    ##############
    # End_Sessions
    ##############
    ####################
    # Start_Publications
    ####################
    $Publication_TB.Add_KeyDown({
		    param ($sender,
			    $e)
		    if ($e.Key -eq [System.Windows.Input.Key]::Enter) { Search_Publication }
	    })
    $Search_Publication_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for a publication.") })
    $Search_Publications_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for all publications in the farm selected.") })
    $Search_Publication.add_Click({ Search_Publication })
    $Publication_settings.add_Click({
		    try
		    {
			    Publications_collapse
			    $listbox_desktop_tag.Items.Clear()
			    $datagrid_application_settings.ItemsSource = $null
			    $datagrid_application_settings_2.ItemsSource = $null
			    $datagrid_desktop_settings_2.ItemsSource = $null
			    $datagrid_desktop_settings_2.ItemsSource = $null
			    Publication_settings
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publication_settings " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Application_settings_Apply.add_Click({
		    try
		    {
			    if ($datagrid_publications.itemsSource.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please search a publication." }
			    elseif ($datagrid_publications.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    else
			    {
				    $Farm = $datagrid_publications.selecteditem.Farm
				    $DDC = ($SyncHash.$Farm).DDC
				    $UID = $datagrid_publications.selecteditem.UID
				    $Type = $datagrid_publications.selecteditem.Type
				    $Name = $datagrid_publications.selecteditem.Name
				    $App = Get-BrokerApplication -AdminAddress $DDC -Uid $UID
				    $App_UUID = $App.UUID
				    Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -PublishedName $datagrid_application_settings.itemssource.Column2Data[3]
				    Get-BrokerApplication -AdminAddress $DDC -UUID $App_UUID | Rename-BrokerApplication -AdminAddress $DDC -NewName $datagrid_application_settings.itemssource.Column2Data[4]
				    $datagrid_publications.Dispatcher.Invoke([Action]{ $datagrid_publications.selecteditem.Name = $datagrid_application_settings.itemssource.Column2Data[4] }, [Windows.Threading.DispatcherPriority]::Normal)
				    Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -BrowserName $datagrid_application_settings.itemssource.Column2Data[5]
				    Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -CommandLineExecutable $datagrid_application_settings.itemssource.Column2Data[6]
				    Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -CommandLineArguments $datagrid_application_settings.itemssource.Column2Data[7]
				    Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Description $datagrid_application_settings.itemssource.Column2Data[8]
				    $datagrid_publications.Dispatcher.Invoke([Action]{ $datagrid_publications.selecteditem.Description = $datagrid_application_settings.itemssource.Column2Data[8] }, [Windows.Threading.DispatcherPriority]::Normal)
				    Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -WorkingDirectory $datagrid_application_settings.itemssource.Column2Data[9]
				    Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -ClientFolder $datagrid_application_settings.itemssource.Column2Data[10]
				    if ($datagrid_application_settings_2.itemssource.Column2SelectedValue[0] -eq $True) { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Enabled $True }
				    else { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Enabled $False }
				    if ($datagrid_application_settings_2.itemssource.Column2SelectedValue[1] -eq $True) { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Visible $True }
				    else { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Visible $False }
				    $Form.Dispatcher.Invoke([Action]{ $datagrid_publications.Items.Refresh() })
				    Show-Dialog_Main -Foreground "Blue" -Text "Changes applied"
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Application_settings_Apply " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Application_settings_Discard.add_Click({
		    try
		    {
			    $datagrid_application_settings.ItemsSource = $null
			    $datagrid_application_settings_2.ItemsSource = $null
			    Publication_settings
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Application_settings_Discard " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Desktop_settings_Apply.add_Click({
		    try
		    {
			    if ($datagrid_publications.itemsSource.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please search a publication." }
			    elseif ($datagrid_publications.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    else
			    {
				    $Farm = $datagrid_publications.selecteditem.Farm
				    $DDC = ($SyncHash.$Farm).DDC
				    $UID = $datagrid_publications.selecteditem.UID
				    $Type = $datagrid_publications.selecteditem.Type
				    $Name = $datagrid_publications.selecteditem.Name
				    $App = Get-BrokerEntitlementPolicyRule -MaxRecordCount 9999 -AdminAddress $DDC -Name $Name -ErrorAction SilentlyContinue
				    $App_UUID = $App.UUID
				    Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -PublishedName $datagrid_desktop_settings.itemssource.Column2Data[4]
				    Rename-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -NewName $datagrid_desktop_settings.itemssource.Column2Data[3]
				    $datagrid_publications.Dispatcher.Invoke([Action]{ $datagrid_publications.selecteditem.Name = $datagrid_desktop_settings.itemssource.Column2Data[3] }, [Windows.Threading.DispatcherPriority]::Normal)
				    Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -Description $datagrid_desktop_settings.itemssource.Column2Data[5]
				    $datagrid_publications.Dispatcher.Invoke([Action]{ $datagrid_publications.selecteditem.Description = $datagrid_desktop_settings.itemssource.Column2Data[5] }, [Windows.Threading.DispatcherPriority]::Normal)
				    if ($datagrid_desktop_settings_2.itemssource.Column2SelectedValue[0] -eq $True) { Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -Enabled $True }
				    else { Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -Enabled $False }
				    $Form.Dispatcher.Invoke([Action]{ $datagrid_publications.Items.Refresh() })
				    Show-Dialog_Main -Foreground "Blue" -Text "Changes applied"
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Desktop_settings_Apply " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Desktop_settings_Discard.add_Click({
		    try
		    {
			    $datagrid_desktop_settings.ItemsSource = $null
			    $datagrid_desktop_settings_2.ItemsSource = $null
			    $listbox_desktop_tag.ItemsSource = $null
			    $listbox_desktop_tag.Items.Clear()
			    Publication_settings
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Desktop_settings_Discard " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $desktop_tag.add_Click({
		    try
		    {
			    $Farm = $datagrid_publications.selecteditem.Farm
			    $DDC = ($SyncHash.$Farm).DDC
			    $UID = $datagrid_publications.selecteditem.UID
			    If ($listbox_desktop_tag.SelectedItem -eq $Null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a TAG to add or change." }
			    else
			    {
				    Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -RestrictToTag $listbox_desktop_tag.SelectedItem
				    Show-Dialog_Main -Foreground "Blue" -Text "TAG modified.`r`nPlease refresh."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_desktop_tag " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $desktop_tag_remove.add_Click({
		    try
		    {
			    $Farm = $datagrid_publications.selecteditem.Farm
			    $DDC = ($SyncHash.$Farm).DDC
			    $UID = $datagrid_publications.selecteditem.UID
			    if ($datagrid_desktop_settings.itemssource.Column2Data[6] -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "No TAG to remove." }
			    else
			    {
				    Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -RestrictToTag $Null
				    Show-Dialog_Main -Foreground "Blue" -Text "TAG removed.`r`nPlease refresh."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_desktop_tag_remove " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Publication_sessions.add_Click({
		    try
		    {
			    $datagrid_PubliSessions.ItemsSource = $null
			    Publications_collapse
			    if ($datagrid_publications.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    else
			    {
				    $Load_TB.Text = "Searching sessions"
				    Refresh_PubliSession
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publication_sessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Kill_PubliSession.add_Click({
		    try
		    {
			    $Sessions = $datagrid_PubliSessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    Else
			    {
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = ($SyncHash.$Session.Farm).DDC
					    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Stop-BrokerSession
				    }
				    Show-Dialog_Main -Foreground "Blue" -Text "Kill command sent.`r`nPlease refresh sessions in few seconds."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Kill_PubliSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Hide_PubliSession.add_Click({
		    try
		    {
			    $Sessions = $datagrid_PubliSessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    else
			    {
				    $VDI = 0
				    $i = 0
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = $SyncHash.($Session.Farm).DDC
					    $Hidden = $Session.Hidden
					    $Type = $Session.Type
					    if ($Type -match "VDI") { $VDI += 1 }
					    else
					    {
						    if ($Hidden -eq $false)
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$true
						    }
						    else
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$false
						    }
					    }
				    }
				    if ($VDI -ne 0 -and $i -ne 0)
				    {
					    $Load_TB.Text = "Refreshing sessions"
					    Refresh_PubliSession
					    $MainLayer.IsEnabled = $false
					    $Main_MB.Foreground = "Blue"
					    $Main_MB.FontSize = "20"
					    $Main_MB.text = "Hide status changed for server connections."
					    $Main_MB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = "`r`nBut hide status for VDI can't be changed."; Foreground = "Red" }))
					    $Dialog_Main.IsOpen = $True
					    $Main_MB_Close.add_Click({
							    $Dialog_Main.IsOpen = $False
							    $MainLayer.IsEnabled = $true
						    })
				    }
				    if ($VDI -ne 0 -and $i -eq 0) { Show-Dialog_Main -Foreground "Red" -Text "Hide status for VDI can't be changed." }
				    if ($VDI -eq 0 -and $i -ne 0)
				    {
					    $Load_TB.Text = "Refreshing sessions"
					    Refresh_PubliSession
					    Show-Dialog_Main -Foreground "Blue" -Text "Hide status changed."
				    }
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Hide_PubliSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Shadow_PubliSession.add_Click({
		    try
		    {
			    $Session = $datagrid_PubliSessions.SelectedItems
			    if ($Session.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    elseif ($Session.count -ge "2") { Show-Dialog_Main -Foreground "Red" -Text "Please select only one session." }
			    else
			    {
				    $Machine = $Session."Machine Name"
				    $User = $Session.User
				    $Domain = $Session.Domain
				    $ID = Get-UserNameSessionIDMap -Comp $Machine | ? { $_.UserName -match $User } | Select-Object -ExpandProperty SessionID
				    $Arg = "/offerra $Machine "
				    $Arg += "$Domain\$User"
				    $Arg += ":"
				    $Arg += $ID
				    Start-Process msra $Arg
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Shadow_PubliSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_PubliSession.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing sessions"
			    Refresh_PubliSession
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_PubliSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_PubliSession.add_Click({
		    try
		    {
			    if ($datagrid_PubliSessions.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No sessions to export." }
			    else
			    {
				    $Name = $datagrid_publications.selecteditem.Name
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportPubliSessions = [hashtable]::Synchronized(@{
						    Name	   = $Name
						    ConfigPath = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportPubliSessions", $SyncHash_ExportPubliSessions)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $Name = $SyncHash_ExportPubliSessions.Name
							    $date = get-date -Format MM_dd_yyyy
							    $Export_PubliSessions = $SyncHash_ExportPubliSessions.ConfigPath + "\Exports\Sessions_$Name-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_PubliSessions)
							    {
								    $i++
								    $Export_PubliSessions = $SyncHash_ExportPubliSessions.ConfigPath + "\Exports\Sessions_$Name-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_PubliSessions.ItemsSource | Export-xlsx -Path $Export_PubliSessions -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_PubliSessions"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_PubliSession_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_PubliSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_PubliSession_Simple.add_Click({
		    $Grid_PubliSessions_Simple.Visibility = "visible"
		    $Grid_PubliSessions_Full.Visibility = "collapse"
	    })
    $Switch_PubliSessions_Full.add_Click({
		    $Grid_PubliSessions_Simple.Visibility = "collapse"
		    $Grid_PubliSessions_Full.Visibility = "visible"
	    })
    $Publication_servers.add_Click({
		    try
		    {
			    $TextBox_Servers_Publications.text = ""
			    $TextBox_TotalServers_Publications.text = ""
			    Publications_collapse
			    if ($datagrid_publications.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    else
			    {
				    $Farm = $datagrid_publications.selecteditem.Farm
				    $DDC = ($SyncHash.$Farm).DDC
				    $UID = $datagrid_publications.selecteditem.UID
				    $Type = $datagrid_publications.selecteditem.Type
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $TextBox_Servers_Publications.Visibility = "Visible"
				    $TextBox_TotalServers_Publications.Visibility = "Visible"
				    $Border_Servers_Publication.Visibility = "Visible"
				    $Load_TB.Text = "Searching servers"
				    $Global:SyncHash_Publiservers_list = [hashtable]::Synchronized(@{
						    DDC  = $DDC
						    UID  = $UID
						    Type = $Type
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_Publiservers_list", $SyncHash_Publiservers_list)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    asnp Citrix*
							    $App = $null
							    $DG_List = @()
							    $AG_List = @()
							    $App_DG_Name = @()
							    $App_AG_Name = @()
							    $App_Serveurs_DG = @()
							    $App_Serveurs_AG = @()
							    $AG_ServersList_1 = @()
							    $AG_ServersList_2 = @()
							    $AG_ServerList = @()
							    $Total_Servers = $null
							    if ($SyncHash_Publiservers_list.Type -eq "Application")
							    {
								    $App = Get-BrokerApplication -AdminAddress $SyncHash_Publiservers_list.DDC -Uid $SyncHash_Publiservers_list.UID
								    $DG_List = Get-BrokerDesktopGroup -AdminAddress $SyncHash_Publiservers_list.DDC | Select-Object name, uid
								    $AG_List = Get-BrokerApplicationGroup -AdminAddress $SyncHash_Publiservers_list.DDC | Select-Object name, uid
								    Foreach ($DG in $DG_List)
								    {
									    Foreach ($App_DG in $App.AssociatedDesktopGroupUids)
									    {
										    if ($App_DG -eq $DG.uid) { $App_DG_Name += $DG.Name }
									    }
								    }
								    Foreach ($AG in $AG_List)
								    {
									    Foreach ($App_AG in $App.AssociatedApplicationGroupUids)
									    {
										    if ($App_AG -eq $AG.uid) { $App_AG_Name += $AG.Name }
									    }
								    }
								    foreach ($DG_Name in $App_DG_Name) { $App_Serveurs_DG += Get-BrokerDesktop -MaxRecordCount 9999 -AdminAddress $SyncHash_Publiservers_list.DDC -DesktopGroupName $DG_Name | Select-Object @{ n = "MachineName"; e = { $_.MachineName.Split('\')[-1] } } }
								    $App_Serveurs_DG = $App_Serveurs_DG.MachineName
								    foreach ($AG_Name in $App_AG_Name) { $App_Serveurs_AG += Get-BrokerApplicationGroup -MaxRecordCount 9999 -AdminAddress $SyncHash_Publiservers_list.DDC -Name $AG_Name }
								    $AG_Servers_1 = $App_Serveurs_AG | Where-Object { $_.RestrictToTag -eq $Null }
								    foreach ($AssociatedDesktopGroupUid in $AG_Servers_1.AssociatedDesktopGroupUids) { $AG_ServersList_1 += Get-BrokerDesktop -AdminAddress $SyncHash_Publiservers_list.DDC -DesktopGroupUid $AssociatedDesktopGroupUid | Select-Object -ExpandProperty @{ n = "MachineName"; e = { $_.MachineName.Split('\')[-1] } } }
								    $AG_ServersList_1 = $AG_ServersList_1.MachineName
								    $AG_Servers_2 = $App_Serveurs_AG | Where-Object { $_.RestrictToTag -ne $Null }
								    foreach ($RestrictToTag in $AG_Servers_2.RestrictToTag) { $AG_ServersList_2 += Get-BrokerDesktop -AdminAddress $SyncHash_Publiservers_list.DDC -Tag $RestrictToTag | Select-Object @{ n = "MachineName"; e = { $_.MachineName.Split('\')[-1] } } }
								    $AG_ServersList_2 = $AG_ServersList_2.MachineName
								    $AG_ServerList = $AG_ServersList_1 + $AG_ServersList_2
								    $Total_Servers = $App_Serveurs_DG + $AG_ServerList
								    $Total_Servers = $Total_Servers | Sort-Object -Unique
								    $Total_Servers_String = [string]::Join([Environment]::NewLine, $Total_Servers)
							    }
							    else
							    {
								    $App = Get-BrokerEntitlementPolicyRule -AdminAddress $SyncHash_Publiservers_list.DDC -Uid $SyncHash_Publiservers_list.UID -ErrorAction SilentlyContinue
								    $DG_List = Get-BrokerDesktopGroup -AdminAddress $SyncHash_Publiservers_list.DDC | Select-Object name, uid
								    Foreach ($DG in $DG_List) { Foreach ($App_DG in $App.DesktopGroupUid) { if ($App_DG -eq $DG.uid) { $DG_Desk = $DG.Name } } }
								    $DG_Servers = Get-BrokerMachine -MaxRecordCount 9999 -AdminAddress $SyncHash_Publiservers_list.DDC | ? { $_.DesktopGroupName -match $DG_Desk } | Select-Object @{ n = "MachineName"; e = { $_.MachineName.Split('\')[-1] } }
								    $DG_Servers = $DG_Servers.MachineName
								    $TAG_Servers = Get-BrokerMachine -MaxRecordCount 9999 -AdminAddress $SyncHash_Publiservers_list.DDC | ? { $_.Tags -match $App.RestrictToTag -and $_.DesktopGroupName -match $DG_Desk } | Select-Object @{ n = "MachineName"; e = { $_.MachineName.Split('\')[-1] } }
								    $TAG_Servers = $TAG_Servers.MachineName
								    If ($App.RestrictToTag -eq $null)
								    {
									    $Total_Servers = $DG_Servers
									    $Total_Servers_String = [string]::Join([Environment]::NewLine, $Total_Servers)
								    }
								    Else
								    {
									    $Total_Servers = $TAG_Servers
									    $Total_Servers_String = [string]::Join([Environment]::NewLine, $Total_Servers)
								    }
							    }
							    if ($Total_Servers.count -eq 0)
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.MainLayer.IsEnabled = $false
										    $SyncHash.Main_MB.Foreground = "Red"
										    $SyncHash.Main_MB.FontSize = "20"
										    $SyncHash.Main_MB.text = "No server configured."
										    $SyncHash.Dialog_Main.IsOpen = $True
									    }, "Normal")
							    }
							    else
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.MainLayer.IsEnabled = $true
									    }, "Normal")
								    $SyncHash.TextBox_Servers_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Publications.Text = $Total_Servers_String }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_TotalServers_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Publications.Text = "Total : " + $Total_Servers.count }, [Windows.Threading.DispatcherPriority]::Normal)
							    }
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publication_servers_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
			    }
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publication_servers " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Publication_access.add_Click({
		    try
		    {
			    $TextBox_Access_Publications.text = ""
			    $TextBox_TotalAccess_Publications.text = ""
			    Publications_collapse
			    if ($datagrid_publications.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    else
			    {
				    $Farm = $datagrid_publications.selecteditem.Farm
				    $DDC = ($SyncHash.$Farm).DDC
				    $UID = $datagrid_publications.selecteditem.UID
				    $Type = $datagrid_publications.selecteditem.Type
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $TextBox_Access_Publications.Visibility = "Visible"
				    $TextBox_TotalAccess_Publications.Visibility = "Visible"
				    $Border_Access_Publication.Visibility = "Visible"
				    $Load_TB.Text = "Searching users access"
				    $Global:SyncHash_Publiaccess_list = [hashtable]::Synchronized(@{
						    DDC  = $DDC
						    UID  = $UID
						    Type = $Type
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_Publiaccess_list", $SyncHash_Publiaccess_list)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    asnp Citrix*
							    $App = $null
							    $DG_List = @()
							    $AG_List = @()
							    $App_DG_Name = @()
							    $App_AG_Name = @()
							    $UserAppList = $null
							    if ($SyncHash_Publiaccess_list.Type -eq "Application")
							    {
								    $App = Get-BrokerApplication -AdminAddress $SyncHash_Publiaccess_list.DDC -Uid $SyncHash_Publiaccess_list.UID
								    $DG_List = Get-BrokerDesktopGroup -AdminAddress $SyncHash_Publiaccess_list.DDC | Select-Object name, uid
								    $AG_List = Get-BrokerApplicationGroup -AdminAddress $SyncHash_Publiaccess_list.DDC | Select-Object name, uid
								    $Users_App = $App.AssociatedUserNames
								    Foreach ($DG in $DG_List)
								    {
									    Foreach ($App_DG in $App.AssociatedDesktopGroupUids)
									    {
										    if ($App_DG -eq $DG.uid) { $App_DG_Name += $DG.Name }
									    }
								    }
								    Foreach ($AG in $AG_List)
								    {
									    Foreach ($App_AG in $App.AssociatedApplicationGroupUids)
									    {
										    if ($App_AG -eq $AG.uid) { $App_AG_Name += $AG.Name }
									    }
								    }
								    if ($App_AG_Name.count -ne 0)
								    {
									    ForEach ($AG in $App_AG_Name)
									    {
										    $AG2 = Get-BrokerApplicationGroup -AdminAddress $SyncHash_Publiaccess_list.DDC -Name $AG
										    foreach ($AssociatedDesktopGroupUid in $AG2.AssociatedDesktopGroupUids)
										    {
											    $DG_Access = Get-BrokerAccessPolicyRule -AdminAddress $SyncHash_Publiaccess_list.DDC -DesktopGroupUid $AssociatedDesktopGroupUid | ? { $_.name -match "Direct" }
											    if ($DG_Access.AllowedUsers -eq "Filtered")
											    {
												    $DG_AG_Access_List = ((Get-BrokerAccessPolicyRule -AdminAddress $SyncHash_Publiaccess_list.DDC -DesktopGroupName $DG_Access.DesktopGroupName | ? { $_.name -match "Direct" } | Select-Object -ExpandProperty IncludedUsers).Name) -split ("`r`n")
												    If ($App.UserFilterEnabled -match "True") { $UserAppList = (($DG_AG_Access_List | ? { $Users_App -contains $_ }) -split ("`r`n")) + $UserAppList }
												    Else { $UserAppList = ($DG_AG_Access_List -split ("`r`n")) + $UserAppList }
											    }
											    Elseif ($AG2.UserFilterEnabled -match "True")
											    {
												    $AG_Access_List = Get-BrokerApplicationGroup -AdminAddress $SyncHash_Publiaccess_list.DDC -Name $AG2.Name | Select-Object -ExpandProperty AssociatedUserNames
												    If ($App.UserFilterEnabled -match "True") { $UserAppList = (($AG_Access_List | ?{ $Users_App -contains $_ }) -split ("`r`n")) + $UserAppList }
												    Else { $UserAppList = ($AG_Access_List -split ("`r`n")) + $UserAppList }
											    }
											    Elseif ($App.UserFilterEnabled -match "True") { $UserAppList = ($Users_App -split ("`r`n")) + $UserAppList }
											    Else { $UserAppList = "Everyone" }
										    }
									    }
								    }
								    if ($App_DG_Name.Count -ne 0)
								    {
									    ForEach ($DGp in $App_DG_Name)
									    {
										    $DG2 = Get-BrokerAccessPolicyRule -AdminAddress $SyncHash_Publiaccess_list.DDC -DesktopGroupName $DGp | ? { $_.name -match "Direct" }
										    if ($DG2.AllowedUsers -match "Filtered")
										    {
											    $DG_Access_List = ((Get-BrokerAccessPolicyRule -AdminAddress $SyncHash_Publiaccess_list.DDC -DesktopGroupName $DG2.DesktopGroupName | ? { $_.name -match "Direct" } | Select-Object -ExpandProperty IncludedUsers).Name)
											    If ($App.UserFilterEnabled -match "True") { $UserAppList += ($DG_Access_List | ?{ $Users_App -contains $_ }) -split ("`r`n") }
											    Else { $UserAppList = ($DG_Access_List -split ("`r`n")) + $UserAppList }
											    Foreach ($item in $DG_Access_List) { if ($item -match "domain") { $UserAppList += $Users_App | ? { $_ -match $item.Split('\')[0] } } }
										    }
										    Elseif ($App.UserFilterEnabled -match "True") { $UserAppList = ($Users_App -split ("`r`n")) + $UserAppList }
										    Else { $UserAppList = "Everyone" }
									    }
								    }
								    $UserAppList = $UserAppList | Sort-Object -Unique
								    $UserAppList_String = [string]::Join([Environment]::NewLine, $UserAppList)
							    }
							    else
							    {
								    $App = Get-BrokerEntitlementPolicyRule -AdminAddress $SyncHash_Publiaccess_list.DDC -Uid $SyncHash_Publiaccess_list.UID -ErrorAction SilentlyContinue
								    $DG_Access = Get-BrokerAccessPolicyRule -AdminAddress $SyncHash_Publiaccess_list.DDC -DesktopGroupUid $App.DesktopGroupUid | ? { $_.name -match "Direct" }
								    If ($App.IncludedUserFilterEnabled -match "False" -and $DG_Access.AllowedUsers -eq "AnyAuthenticated") { $UserAppList = "Everyone" }
								    If ($App.IncludedUserFilterEnabled -match "True" -and $DG_Access.AllowedUsers -eq "AnyAuthenticated") { $UserAppList = $App.IncludedUsers.Name }
								    If ($App.IncludedUserFilterEnabled -match "False" -and $DG_Access.AllowedUsers -eq "Filtered") { $UserAppList = $DG_Access.IncludedUsers.Name }
								    If ($App.IncludedUserFilterEnabled -match "True" -and $DG_Access.AllowedUsers -eq "Filtered") { $UserAppList = (($DG_Access.IncludedUsers.Name | ?{ $App.IncludedUsers.Name -contains $_ }) -split ("`r`n")) }
								    $UserAppList_String = [string]::Join([Environment]::NewLine, $UserAppList)
							    }
							    if ($UserAppList.count -eq 0)
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.MainLayer.IsEnabled = $false
										    $SyncHash.Main_MB.Foreground = "Red"
										    $SyncHash.Main_MB.FontSize = "20"
										    $SyncHash.Main_MB.text = "No user configured."
										    $SyncHash.Dialog_Main.IsOpen = $True
									    }, "Normal")
							    }
							    else
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.MainLayer.IsEnabled = $true
									    }, "Normal")
								    $SyncHash.TextBox_Access_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Access_Publications.Text = $UserAppList_String }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_TotalAccess_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalAccess_Publications.Text = "Total : " + $UserAppList.count }, [Windows.Threading.DispatcherPriority]::Normal)
							    }
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publication_access_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
			    }
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Publication_access " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $S_Publications_Details.Add_SelectionChanged({
		    try
		    {
			    if ($S_Publications_Details.selectedItem -ne $null)
			    {
				    $Publication_TB.text = ""
				    $datagrid_publications.Visibility = "Collapsed"
				    $Publication_settings.Visibility = "Collapsed"
				    $Publication_sessions.Visibility = "Collapsed"
				    $Publication_servers.Visibility = "Collapsed"
				    $Publication_access.Visibility = "Collapsed"
				    $Load_TB.Text = "Searching publications"
				    Publications_collapse
				    $TB_AllPublis.Visibility = "Visible"
				    $Border_AllPublis.Visibility = "Visible"
				    Refresh_AllPublis
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_Publications_Details " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Disable_AllPublis.add_Click({
		    try
		    {
			    $Publis = $datagrid_AllPublications.SelectedItems
			    if ($Publis.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    Else
			    {
				    foreach ($Publi in $Publis)
				    {
					    $UID = $Publi.Uid
					    $Enabled = $Publi.Enabled
					    $DDC = ($SyncHash.$Publi.Farm).DDC
					    if ($Enabled -eq $false) { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Enabled $true }
					    else { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Enabled $false }
				    }
				    $Load_TB.Text = "Refreshing publications"
				    Refresh_AllPublis
				    Show-Dialog_Main -Foreground "Blue" -Text "Enabled status changed for the selected publications."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Disable_AllPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Hide_AllPublis.add_Click({
		    try
		    {
			    $Publis = $datagrid_AllPublications.SelectedItems
			    if ($Publis.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    Else
			    {
				    foreach ($Publi in $Publis)
				    {
					    $UID = $Publi.Uid
					    $Visible = $Publi.Visible
					    $DDC = ($SyncHash.$Publi.Farm).DDC
					    if ($Visible -eq $false) { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Visible $true }
					    else { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Visible $false }
				    }
				    $Load_TB.Text = "Refreshing publications"
				    Refresh_AllPublis
				    Show-Dialog_Main -Foreground "Blue" -Text "Visible state changed for the selected publications."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Hide_AllPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Delete_AllPublis.add_Click({
		    try
		    {
			    $Publis = $datagrid_AllPublications.SelectedItems
			    if ($Publis.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to delete the selected publications ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $Main_MB_Confirm.add_Click({
						    $Publis = $datagrid_AllPublications.SelectedItems
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
						    foreach ($Publi in $Publis)
						    {
							    $UID = $Publi.Uid
							    $Type = $Publi.Type
							    $DDC = ($SyncHash.$Publi.Farm).DDC
							    if ($Type -eq "Application") { Remove-BrokerApplication -AdminAddress $DDC -InputObject $UID }
							    else { Remove-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID }
						    }
						    $Load_TB.Text = "Refreshing publications"
						    Refresh_AllPublis
						    Show-Dialog_Main -Foreground "Blue" -Text "Selected publications deleted."
					    })
				    $Main_MB_Cancel.add_Click({
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Delete_AllPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_AllPublis.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing publications"
			    Refresh_AllPublis
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_AllPublis.add_Click({
		    try
		    {
			    if ($datagrid_AllPublications.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No publications to export." }
			    else
			    {
				    $FarmSelected = $S_Publications_Details.selectedItem
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportPublications = [hashtable]::Synchronized(@{
						    FarmSelected = $FarmSelected
						    ConfigPath   = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportPublications", $SyncHash_ExportPublications)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $Farm = $SyncHash_ExportPublications.FarmSelected
							    $date = get-date -Format MM_dd_yyyy
							    $Export_Publications = $SyncHash_ExportPublications.ConfigPath + "\Exports\Publications_$Farm-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_Publications)
							    {
								    $i++
								    $Export_Publications = $SyncHash_ExportPublications.ConfigPath + "\Exports\Publications_$Farm-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_AllPublications.ItemsSource | Export-xlsx -Path $Export_Publications -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_Publications"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_AllPublis_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_AllPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_AllPublis_Simple.add_Click({
		    $Grid_AllPublications_Simple.Visibility = "visible"
		    $Grid_AllPublications_Full.Visibility = "collapse"
	    })
    $Switch_AllPublis_Full.add_Click({
		    $Grid_AllPublications_Simple.Visibility = "collapse"
		    $Grid_AllPublications_Full.Visibility = "visible"
	    })
    ##################
    # End_Publications
    ##################
    #############
    # Start_VDAs
    #############
    $VDA_TB.Add_KeyDown({
		    param ($sender,
			    $e)
		    if ($e.Key -eq [System.Windows.Input.Key]::Enter) { Search_VDA }
	    })
    $Search_VDA_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for a VDA.") })
    $Search_AllVDAs_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for all VDAs in the farm selected.") })
    $Search_VDA.add_Click({ Search_VDA })
    $VDA_Details.add_Click({
		    try
		    {
			    VDAs_collapse
			    if ($datagrid_VDAsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    else
			    {
				    $datagrid_VDA_settings.ItemsSource = $null
				    $Farm = $datagrid_VDAsList.selecteditem.farm
				    $UID = $datagrid_VDAsList.selecteditem.uid
				    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
				    $DDC = ($SyncHash.$Farm).DDC
				    $VDA = Get-BrokerMachine -AdminAddress $DDC -Uid $UID
				    $VDA_Registration = $VDA.RegistrationState
				    If ($VDA.InMaintenanceMode -eq "True") { $VDA_Maintenance = "Enable" }
				    else { $VDA_Maintenance = "Disable" }
				    $VDA_Power = $VDA.PowerState
				    $VDA_Farm = $Farm
				    $VDA_OS = $VDA.OSType
				    $VDA_IP = $VDA.IPAddress
				    $VDA_DG = $VDA.DesktopGroupName
				    $VDA_MC = $VDA.CatalogName
				    $VDA_Load = $VDA.LoadIndex
				    $VDA_Agent = $VDA.AgentVersion
				    $VDA_Provisioning = $VDA.ProvisioningType
				    $Applications_Groups = @()
				    if ($VDA.PowerState -eq "On")
				    {
					    if (Test-Connection -Count 1 -quiet $Name) { $VDA_BootTime = Get-CimInstance -ComputerName ([System.Net.Dns]::GetHostByName($Name)).HostName -ClassName Win32_OperatingSystem | Select -ExpandProperty LastBootUpTime }
					    else { $VDA_BootTime = "N/A" }
				    }
				    else { $VDA_BootTime = "N/A" }
				    If (($VDA.Tags).count -ne 0)
				    {
					    $VDA_TAG = $VDA.Tags | Out-String
					    foreach ($Tag in $VDA.Tags) { $Applications_Groups += Get-BrokerApplicationGroup -AdminAddress VWC2APP141 -RestrictToTag $Tag | Select-Object Name }
				    }
				    else { $VDA_TAG = "No Tag" }
				    if ($Applications_Groups.Name.count -eq "0") { $VDA_AG = "0" }
				    else { $VDA_AG = $Applications_Groups.Name | Out-String }
				    $datagrid_VDA_settings.Visibility = "Visible"
				    $Enable_Maintenance_VDA.Visibility = "Visible"
				    $Disable_Maintenance_VDA.Visibility = "Visible"
				    $PowerOn_VDA.Visibility = "Visible"
				    $PowerOff_VDA.Visibility = "Visible"
				    $Refresh_VDA.Visibility = "Visible"
				    $datagrid_VDA_settings.ItemsSource = @(
					    [PSCustomObject]@{ Column1Header = "Registration State"; Column2Data = $VDA_Registration; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Maintenance State"; Column2Data = $VDA_Maintenance; IsReadOnly = $true }
					    [PSCustomObject]@{ Column1Header = "Power State"; Column2Data = $VDA_Power; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Farm"; Column2Data = $VDA_Farm; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "OS Type"; Column2Data = $VDA_OS; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "IP Address"; Column2Data = $VDA_IP; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Delevery Group"; Column2Data = $VDA_DG; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Machine Catalog"; Column2Data = $VDA_MC; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Load"; Column2Data = $VDA_Load; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Agent Version"; Column2Data = $VDA_Agent; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Provisioning Type"; Column2Data = $VDA_Provisioning; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Last Boot Time"; Column2Data = $VDA_BootTime; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Tags"; Column2Data = $VDA_TAG; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Applications Groups"; Column2Data = $VDA_AG; IsReadOnly = $true }
				    )
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDA_Details " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Enable_Maintenance_VDA.add_Click({
		    try
		    {
			    $Farm = $datagrid_VDAsList.selecteditem.farm
			    $UID = $datagrid_VDAsList.selecteditem.uid
			    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
			    $DDC = ($SyncHash.$Farm).DDC
			    $VDA = Get-BrokerMachine -AdminAddress $DDC -Uid $UID
			    If ($VDA.InMaintenanceMode -eq $True) { Show-Dialog_Main -Foreground "Red" -Text "$Name already in maintenance." }
			    else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to enable maintenance for $Name ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $Main_MB_Confirm.add_Click({
						    $Farm = $datagrid_VDAsList.selecteditem.farm
						    $UID = $datagrid_VDAsList.selecteditem.uid
						    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
						    $DDC = ($SyncHash.$Farm).DDC
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
						    Set-BrokerMachineMaintenanceMode -AdminAddress $DDC -InputObject $UID $true
						    Refresh_VDA
						    Show-Dialog_Main -Foreground "Blue" -Text "Maintenance enabled for $Name."
					    })
				    $Main_MB_Cancel.add_Click({
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Enable_Maintenance_VDA " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Disable_Maintenance_VDA.add_Click({
		    try
		    {
			    $Farm = $datagrid_VDAsList.selecteditem.farm
			    $UID = $datagrid_VDAsList.selecteditem.uid
			    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
			    $DDC = ($SyncHash.$Farm).DDC
			    $VDA = Get-BrokerMachine -AdminAddress $DDC -Uid $UID
			    If ($VDA.InMaintenanceMode -eq $False) { Show-Dialog_Main -Foreground "Red" -Text "$Name is not in maintenance." }
			    else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to disable maintenance for $Name ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $Main_MB_Confirm.add_Click({
						    $Farm = $datagrid_VDAsList.selecteditem.farm
						    $UID = $datagrid_VDAsList.selecteditem.uid
						    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
						    $DDC = ($SyncHash.$Farm).DDC
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
						    Set-BrokerMachineMaintenanceMode -AdminAddress $DDC -InputObject $UID $False
						    Refresh_VDA
						    Show-Dialog_Main -Foreground "Blue" -Text "Maintenance disabled for $Name."
					    })
				    $Main_MB_Cancel.add_Click({
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Disable_Maintenance_VDA " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOn_VDA.add_Click({
		    try
		    {
			    $Farm = $datagrid_VDAsList.selecteditem.farm
			    $UID = $datagrid_VDAsList.selecteditem.uid
			    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
			    $DDC = ($SyncHash.$Farm).DDC
			    $VDA = Get-BrokerMachine -AdminAddress $DDC -Uid $UID
			    If ($VDA.PowerState -eq "On") { Show-Dialog_Main -Foreground "Red" -Text "$Name is already powered on." }
			    else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power on $Name ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $Main_MB_Confirm.add_Click({
						    $Farm = $datagrid_VDAsList.selecteditem.farm
						    $UID = $datagrid_VDAsList.selecteditem.uid
						    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
						    $DDC = ($SyncHash.$Farm).DDC
						    $VDA = Get-BrokerMachine -AdminAddress $DDC -Uid $UID
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
						    New-BrokerHostingPowerAction -AdminAddress $DDC -MachineName $VDA.MachineName -Action TurnOn
						    Show-Dialog_Main -Foreground "Blue" -Text "Power On action has been sent for $Name.`r`nPlease refresh."
					    })
				    $Main_MB_Cancel.add_Click({
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOn_VDA " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOff_VDA.add_Click({
		    try
		    {
			    $Farm = $datagrid_VDAsList.selecteditem.farm
			    $UID = $datagrid_VDAsList.selecteditem.uid
			    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
			    $DDC = ($SyncHash.$Farm).DDC
			    $VDA = Get-BrokerMachine -AdminAddress $DDC -Uid $UID
			    If ($VDA.PowerState -eq "Off") { Show-Dialog_Main -Foreground "Red" -Text "$Name is already powered off." }
			    else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power off $Name ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $Main_MB_Confirm.add_Click({
						    $Farm = $datagrid_VDAsList.selecteditem.farm
						    $UID = $datagrid_VDAsList.selecteditem.uid
						    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
						    $DDC = ($SyncHash.$Farm).DDC
						    $VDA = Get-BrokerMachine -AdminAddress $DDC -Uid $UID
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
						    New-BrokerHostingPowerAction -AdminAddress $DDC -MachineName $VDA.MachineName -Action TurnOff
						    Show-Dialog_Main -Foreground "Blue" -Text "Power Off action has been sent for $Name.`r`nPlease refresh."
					    })
				    $Main_MB_Cancel.add_Click({
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOff_VDA " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_VDA.add_Click({ Refresh_VDA })
    $VDA_Sessions.add_Click({
		    try
		    {
			    $datagrid_VDA_sessions.ItemsSource = $null
			    VDAs_collapse
			    if ($datagrid_VDAsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    else
			    {
				    $Farm = $datagrid_VDAsList.selecteditem.farm
				    $UID = $datagrid_VDAsList.selecteditem.uid
				    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
				    $DDC = ($SyncHash.$Farm).DDC
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Searching sessions"
				    $Global:SyncHash_VDASessions_list = [hashtable]::Synchronized(@{
						    DDC  = $DDC
						    UID  = $UID
						    Name = $Name
						    Farm = $Farm
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_VDASessions_list", $SyncHash_VDASessions_list)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    asnp Citrix*
							    $VDA_Sessions_List = Get-BrokerSession -MaxRecordCount 99999 -AdminAddress $SyncHash_VDASessions_list.DDC -MachineUid $SyncHash_VDASessions_list.UID | Select-Object @{
								    n = "User"; e = {
									    if ($_.UserFullName -eq $null) { ".no data" }
									    else { $_.UserFullName }
								    }
							    }, @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Farm"; e = { $SyncHash_VDASessions_list.Farm } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, Hidden, @{ n = "Session State"; e = { $_.SessionState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Session Type"; e = { $_.SessionType } }, @{
								    n = "Type"; e = {
									    if ($_.SessionSupport -match "MultiSession") { "Server" }
									    else { "VDI" }
								    }
							    }, Protocol, @{ n = "Start Time"; e = { $_.Starttime } }, @{ n = "Applications"; e = { $_.LaunchedViaPublishedName } }, @{ n = "Client Name"; e = { $_.ClientName } }, @{ n = "Client Address"; e = { $_.ClientAddress } }, @{ n = "DDC"; e = { $_.ControllerDNSName } }, UID
							    $VDA_Sessions_List = $VDA_Sessions_List | Sort-Object User
							    $Total_VDAs_Sessions = $VDA_Sessions_List.User.count
							    $Simple_VDAs_Sessions_List = $VDA_Sessions_List.User
							    if ($Total_VDAs_Sessions -eq 0) { $Simple_VDAs_Sessions_List_String = $null }
							    else { $Simple_VDAs_Sessions_List_String = [string]::Join([Environment]::NewLine, $Simple_VDAs_Sessions_List) }
							    if ($Total_VDAs_Sessions -eq 0)
							    {
								    $SyncHash.TextBox_TotalVDAs_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalVDAs_Sessions.text = "Total = $Total_VDAs_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.MainLayer.IsEnabled = $false
										    $SyncHash.Main_MB.Foreground = "Red"
										    $SyncHash.Main_MB.FontSize = "20"
										    $SyncHash.Main_MB.text = "No session found."
										    $SyncHash.Dialog_Main.IsOpen = $True
									    }, "Normal")
							    }
							    elseif ($Total_VDAs_Sessions -eq 1)
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.Grid_VDAs_Sessions_Full.Visibility = "Visible"
										    $SyncHash.MainLayer.IsEnabled = $true
									    }, "Normal")
								    $VDASessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
								    $VDASessions_List_Datagrid.Add($VDA_Sessions_List)
								    $SyncHash.datagrid_VDA_sessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_VDA_sessions.ItemsSource = $VDASessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_TotalVDAs_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalVDAs_Sessions.text = "Total = $Total_VDAs_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_VDAs_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_VDAs_Sessions.text = $Simple_VDAs_Sessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
							    }
							    else
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.Grid_VDAs_Sessions_Full.Visibility = "Visible"
										    $SyncHash.MainLayer.IsEnabled = $true
									    }, "Normal")
								    $VDASessions_List_Datagrid = New-Object System.Collections.Generic.List[Object]
								    $VDASessions_List_Datagrid.AddRange($VDA_Sessions_List)
								    $SyncHash.datagrid_VDA_sessions.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_VDA_sessions.ItemsSource = $VDASessions_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_TotalVDAs_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalVDAs_Sessions.text = "Total = $Total_VDAs_Sessions" }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_VDAs_Sessions.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_VDAs_Sessions.text = $Simple_VDAs_Sessions_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
							    }
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDA_Sessions_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDA_Sessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Kill_VDASession.add_Click({
		    try
		    {
			    $Sessions = $datagrid_VDA_sessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    Else
			    {
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = ($SyncHash.$Session.Farm).DDC
					    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Stop-BrokerSession
				    }
				    Show-Dialog_Main -Foreground "Blue" -Text "Kill command sent.`r`nPlease refresh sessions in few seconds."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Kill_VDASession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Hide_VDASession.add_Click({
		    try
		    {
			    $Sessions = $datagrid_VDA_sessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    else
			    {
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = $SyncHash.($Session.Farm).DDC
					    $Hidden = $Session.Hidden
					    if ($Hidden -eq $false) { Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$true }
					    else { Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$false }
				    }
				    Refresh_VDASession
				    Show-Dialog_Main -Foreground "Blue" -Text "Hide status changed."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Hide_VDASession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Shadow_VDASession.add_Click({
		    try
		    {
			    $Session = $datagrid_VDA_sessions.SelectedItems
			    if ($Session.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    elseif ($Session.count -ge "2") { Show-Dialog_Main -Foreground "Red" -Text "Please select only one session." }
			    else
			    {
				    $Machine = $Session."Machine Name"
				    $User = $Session.User
				    $Domain = $Session.Domain
				    $ID = Get-UserNameSessionIDMap -Comp $Machine | ? { $_.UserName -match $User } | Select-Object -ExpandProperty SessionID
				    $Arg = "/offerra $Machine "
				    $Arg += "$Domain\$User"
				    $Arg += ":"
				    $Arg += $ID
				    Start-Process msra $Arg
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Shadow_VDASession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_VDASession.add_Click({ Refresh_VDASession })
    $Export_VDASession.add_Click({
		    try
		    {
			    if ($datagrid_VDA_sessions.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No sessions to export." }
			    else
			    {
				    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportVDASessions = [hashtable]::Synchronized(@{
						    Name	   = $Name
						    ConfigPath = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportVDASessions", $SyncHash_ExportVDASessions)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $Name = $SyncHash_ExportVDASessions.Name
							    $date = get-date -Format MM_dd_yyyy
							    $Export_VDASessions = $SyncHash_ExportVDASessions.ConfigPath + "\Exports\Sessions_$Name-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_VDASessions)
							    {
								    $i++
								    $Export_VDASessions = $SyncHash_ExportVDASessions.ConfigPath + "\Exports\Sessions_$Name-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_VDA_sessions.ItemsSource | Export-xlsx -Path $Export_VDASessions -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_VDASessions"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_VDASession_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_VDASession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_VDAs_Sessions_Simple.add_Click({
		    $Grid_VDAs_Sessions_Simple.Visibility = "visible"
		    $Grid_VDAs_Sessions_Full.Visibility = "collapse"
	    })
    $Switch_VDAs_Sessions_Full.add_Click({
		    $Grid_VDAs_Sessions_Simple.Visibility = "collapse"
		    $Grid_VDAs_Sessions_Full.Visibility = "visible"
	    })
    $VDA_Publications.add_Click({
		    try
		    {
			    $TextBox_VDA_Publications.text = ""
			    $TextBox_TotalVDA_Publications.text = ""
			    VDAs_collapse
			    if ($datagrid_VDAsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    else
			    {
				    $Farm = $datagrid_VDAsList.selecteditem.farm
				    $UID = $datagrid_VDAsList.selecteditem.uid
				    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
				    $DDC = ($SyncHash.$Farm).DDC
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Searching publications"
				    $Global:SyncHash_VDAPublications_list = [hashtable]::Synchronized(@{
						    DDC						      = $DDC
						    UID						      = $UID
						    Border_VDA_Publications	      = $Border_VDA_Publications
						    TextBox_VDA_Publications	  = $TextBox_VDA_Publications
						    TextBox_TotalVDA_Publications = $TextBox_TotalVDA_Publications
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_VDAPublications_list", $SyncHash_VDAPublications_list)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    asnp Citrix*
							    $Applications_Groups = @()
							    $Application_Group_UID = @()
							    $VDA_Publications = @()
							    $VDA = Get-BrokerMachine -AdminAddress $SyncHash_VDAPublications_list.DDC -Uid $SyncHash_VDAPublications_list.UID
							    If (($VDA.Tags).count -ne 0)
							    {
								    foreach ($Tag in $VDA.Tags)
								    {
									    $Applications_Groups += Get-BrokerApplicationGroup -AdminAddress $SyncHash_VDAPublications_list.DDC -RestrictToTag $Tag | Select-Object Name
									    $Application_Group_UID += Get-BrokerApplicationGroup -AdminAddress $SyncHash_VDAPublications_list.DDC -RestrictToTag $Tag | Select-Object UID
								    }
								    $VDA_AG = $Applications_Groups.Name | Out-String
								    $VDA_AG_UID = $Application_Group_UID.UID
								    foreach ($AG in $VDA_AG_UID) { $VDA_Publications += Get-BrokerApplication -AdminAddress $SyncHash_VDAPublications_list.DDC -ApplicationGroupUid $AG | Select-Object -ExpandProperty PublishedName }
								    $VDA_Publications += Get-BrokerApplication -AdminAddress $SyncHash_VDAPublications_list.DDC -AssociatedDesktopGroupUid $VDA.desktopgroupUID | Select-Object -ExpandProperty PublishedName
							    }
							    Else
							    {
								    $Application_Groups = Get-BrokerApplicationGroup -AdminAddress $SyncHash_VDAPublications_list.DDC -AssociatedDesktopGroupUid $VDA.desktopgroupUID
								    $VDA_AG = $Application_Groups.Name | Out-String
								    $VDA_AG_UID = $Application_Group_UID.UID
								    foreach ($AG in $VDA_AG_UID) { $VDA_Publications += Get-BrokerApplication -AdminAddress $SyncHash_VDAPublications_list.DDC -ApplicationGroupUid $AG | Select-Object -ExpandProperty PublishedName }
								    $VDA_Publications += Get-BrokerApplication -AdminAddress $SyncHash_VDAPublications_list.DDC -AssociatedDesktopGroupUid $VDA.desktopgroupUID | Select-Object -ExpandProperty PublishedName
							    }
							    $VDA_Publications_String = [string]::Join([Environment]::NewLine, $VDA_Publications)
							    if ($VDA_Publications.count -eq 0)
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.MainLayer.IsEnabled = $false
										    $SyncHash.Main_MB.Foreground = "Red"
										    $SyncHash.Main_MB.FontSize = "20"
										    $SyncHash.Main_MB.text = "No publications."
										    $SyncHash.Dialog_Main.IsOpen = $True
									    }, "Normal")
							    }
							    else
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.Border_VDA_Publications.Visibility = "Visible"
										    $SyncHash.TextBox_VDA_Publications.Visibility = "Visible"
										    $SyncHash.TextBox_TotalVDA_Publications.Visibility = "Visible"
										    $SyncHash.MainLayer.IsEnabled = $true
									    }, "Normal")
								    $SyncHash.TextBox_VDA_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_VDA_Publications.Text = $VDA_Publications_String }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_TotalVDA_Publications.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalVDA_Publications.Text = "Total : " + $VDA_Publications.count }, [Windows.Threading.DispatcherPriority]::Normal)
							    }
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDA_Publications_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
			    }
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDA_Publications " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $VDA_Hotfixes.add_Click({
		    try
		    {
			    $datagrid_VDA_Hotfixes.ItemsSource = $null
			    VDAs_collapse
			    if ($datagrid_VDAsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    else
			    {
				    $Farm = $datagrid_VDAsList.selecteditem.farm
				    $UID = $datagrid_VDAsList.selecteditem.uid
				    $Name = $datagrid_VDAsList.selecteditem.'Machine Name'
				    $DDC = ($SyncHash.$Farm).DDC
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Searching hotfixes"
				    $Global:SyncHash_VDAHotfixes_list = [hashtable]::Synchronized(@{
						    Name					  = $Name
						    Border_VDA_Hotfixes	      = $Border_VDA_Hotfixes
						    datagrid_VDA_Hotfixes	  = $datagrid_VDA_Hotfixes
						    TextBox_TotalVDA_Hotfixes = $TextBox_TotalVDA_Hotfixes
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_VDAHotfixes_list", $SyncHash_VDAHotfixes_list)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $VDA_HotFixList = @()
							    $VDA_HotFixList = Get-HotFix -ComputerName $SyncHash_VDAHotfixes_list.Name | Select-Object HotFixID, @{ n = "Installed On"; e = { $_.InstalledOn } } | Sort-Object HotFixID
							 #   $VDAHotFixList_String = [string]::Join([Environment]::NewLine, $VDA_HotFixList)
							    if ($VDA_HotFixList.HotFixID.count -eq 0)
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.MainLayer.IsEnabled = $false
										    $SyncHash.Main_MB.Foreground = "Red"
										    $SyncHash.Main_MB.FontSize = "20"
										    $SyncHash.Main_MB.text = "No hotfixes found."
										    $SyncHash.Dialog_Main.IsOpen = $True
									    }, "Normal")
							    }
							    elseif ($VDA_HotFixList.HotFixID.count -eq 1)
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.Border_VDA_Hotfixes.Visibility = "Visible"
										    $SyncHash.datagrid_VDA_Hotfixes.Visibility = "Visible"
										    $SyncHash.TextBox_TotalVDA_Hotfixes.Visibility = "Visible"
										    $SyncHash.MainLayer.IsEnabled = $true
									    }, "Normal")
								    $VDA_HotFixList_Datagrid = New-Object System.Collections.Generic.List[Object]
								    $VDA_HotFixList_Datagrid.Add($VDA_HotFixList)
								    $SyncHash.datagrid_VDA_Hotfixes.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_VDA_Hotfixes.ItemsSource = $VDA_HotFixList_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_TotalVDA_Hotfixes.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalVDA_Hotfixes.Text = "Total : " + $VDA_HotFixList.HotFixID.count }, [Windows.Threading.DispatcherPriority]::Normal)
							    }
							    else
							    {
								    $SyncHash.Form.Dispatcher.Invoke([action]{
										    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
										    $SyncHash.Border_VDA_Hotfixes.Visibility = "Visible"
										    $SyncHash.datagrid_VDA_Hotfixes.Visibility = "Visible"
										    $SyncHash.TextBox_TotalVDA_Hotfixes.Visibility = "Visible"
										    $SyncHash.MainLayer.IsEnabled = $true
									    }, "Normal")
								    $VDA_HotFixList_Datagrid = New-Object System.Collections.Generic.List[Object]
								    $VDA_HotFixList_Datagrid.AddRange($VDA_HotFixList)
								    $SyncHash.datagrid_VDA_Hotfixes.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_VDA_Hotfixes.ItemsSource = $VDA_HotFixList_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
								    $SyncHash.TextBox_TotalVDA_Hotfixes.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalVDA_Hotfixes.Text = "Total : " + $VDA_HotFixList.HotFixID.count }, [Windows.Threading.DispatcherPriority]::Normal)
							    }
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDA_Hotfixes_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
			    }
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_VDA_Hotfixes " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $S_VDAs_Details.Add_SelectionChanged({
		    try
		    {
			    if ($S_VDAs_Details.selectedItem -ne $null)
			    {
				    $VDA_TB.text = ""
				    $datagrid_VDAsList.Visibility = "Collapsed"
				    $VDA_Details.Visibility = "Collapsed"
				    $VDA_Sessions.Visibility = "Collapsed"
				    $VDA_Publications.Visibility = "Collapsed"
				    $VDA_Hotfixes.Visibility = "Collapsed"
				    VDAs_collapse
				    $Border_AllVDAs.Visibility = "Visible"
				    $TB_AllVDAs.Visibility = "Visible"
				    $Load_TB.Text = "Searching VDAs"
				    Refresh_AllVDAs
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_VDAs_Details " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Enable_Maintenance_AllVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_AllVDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to enable maintenance for the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "Enable_Maintenance_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_AllVDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Enable_Maintenance_AllVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Disble_Maintenance_AllVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_AllVDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $lastButtonClicked = $null
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to disable maintenance for the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "Disble_Maintenance_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_AllVDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Disble_Maintenance_AllVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOn_AllVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_AllVDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power on the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "PowerOn_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_AllVDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOn_AllVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOff_AllVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_AllVDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power off the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "PowerOff_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_AllVDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOff_AllVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_AllVDAs.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing VDAs"
			    Refresh_AllVDAs
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_AllVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_AllVDAs.add_Click({
		    try
		    {
			    if ($datagrid_AllVDAs.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No VDAs to export." }
			    else
			    {
				    $FarmSelected = $S_VDAs_Details.selecteditem
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportVDAs = [hashtable]::Synchronized(@{
						    FarmSelected = $FarmSelected
						    ConfigPath   = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportVDAs", $SyncHash_ExportVDAs)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $Farm = $SyncHash_ExportVDAs.FarmSelected
							    $date = get-date -Format MM_dd_yyyy
							    $Export_VDAs = $SyncHash_ExportVDAs.ConfigPath + "\Exports\VDAs_$Farm-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_VDAs)
							    {
								    $i++
								    $Export_VDAs = $SyncHash_ExportVDAs.ConfigPath + "\Exports\VDAs_$Farm-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_AllVDAs.ItemsSource | Export-xlsx -Path $Export_VDAs -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_VDAs"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_AllVDAs_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_AllVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_AllVDAs_Simple.add_Click({
		    $Grid_AllVDAs_Simple.Visibility = "visible"
		    $Grid_AllVDAs_Full.Visibility = "collapse"
	    })
    $Switch_AllVDAs_Full.add_Click({
		    $Grid_AllVDAs_Simple.Visibility = "collapse"
		    $Grid_AllVDAs_Full.Visibility = "visible"
	    })
    ###########
    # End_VDAs
    ###########
    #######################
    # Start_MachineCatalogs
    #######################
    $MC_TB.Add_KeyDown({
		    param ($sender,
			    $e)
		    if ($e.Key -eq [System.Windows.Input.Key]::Enter) { Search_MC }
	    })
    $Search_MC_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for a Machine Catalog.") })
    $Search_AllMCs_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for all Machine Catalogs in the farm selected.") })
    $Search_MC.add_Click({
		    $Load_TB.Text = "Searching Machine Catalogs"
		    Search_MC
	    })
    $MC_Details.add_Click({
		    try
		    {
			    MCs_collapse
			    if ($datagrid_MCsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a Machine Catalog." }
			    else
			    {
				    $datagrid_MC_settings.ItemsSource = $null
				    $Farm = $datagrid_MCsList.selecteditem.farm
				    $UID = $datagrid_MCsList.selecteditem.uid
				    $Name = $datagrid_MCsList.selecteditem.name
				    $DDC = ($SyncHash.$Farm).DDC
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Searching Machine Catalog Details"
				    $Global:SyncHash_MC_Details = [hashtable]::Synchronized(@{
						    Farm				 = $Farm
						    UID				     = $UID
						    DDC				     = $DDC
						    Name				 = $Name
						    datagrid_MC_settings = $datagrid_MC_settings
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_MC_Details", $SyncHash_MC_Details)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    asnp Citrix*
							    $MC = Get-BrokerCatalog -AdminAddress $SyncHash_MC_Details.DDC -Uid $SyncHash_MC_Details.UID
							    $MC_VDAs = (Get-BrokerMachine -AdminAddress $SyncHash_MC_Details.DDC -MaxRecordCount 999999 | ? { $_.CatalogName -eq $SyncHash_MC_Details.Name }).count
							    $MC_Name = $SyncHash_MC_Details.Name
							    $MC_Description = $MC.Description
							    $MC_Farm = $SyncHash_MC_Details.Farm
							    $MC_SessionSupport = $MC.SessionSupport
							    $MC_AllocationType = $MC.AllocationType
							    $MC_AvailableCount = $MC.AvailableCount
							    $MC_UsedCount = $MC.UsedCount
							    $MC_MachinesArePhysical = $MC.MachinesArePhysical
							    $MC_PersistUserChanges = $MC.PersistUserChanges
							    $MC_ProvisioningType = $MC.ProvisioningType
							    $SyncHash_MC_Details.datagrid_MC_settings.Dispatcher.Invoke([Action]{ $SyncHash_MC_Details.datagrid_MC_settings.Visibility = "Visible" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash_MC_Details.datagrid_MC_settings.Dispatcher.Invoke([Action]{
									    $SyncHash_MC_Details.datagrid_MC_settings.ItemsSource = @(
										    [PSCustomObject]@{ Column1Header = "Machine Catalog"; Column2Data = $MC_Name; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Description"; Column2Data = $MC_Description; IsReadOnly = $true }
										    [PSCustomObject]@{ Column1Header = "Number of VDAs"; Column2Data = $MC_VDAs; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Farm"; Column2Data = $MC_Farm; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Session Support"; Column2Data = $MC_SessionSupport; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Allocation Type"; Column2Data = $MC_AllocationType; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Available Count"; Column2Data = $MC_AvailableCount; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Used Count"; Column2Data = $MC_UsedCount; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Machines Are Physical"; Column2Data = $MC_MachinesArePhysical; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Provisioning Type"; Column2Data = $MC_ProvisioningType; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Persist User Changes"; Column2Data = $MC_PersistUserChanges; IsReadOnly = $true })
								    }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_MC_Details_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_MC_Details " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $MC_VDAs.add_Click({
		    try
		    {
			    MCs_collapse
			    $datagrid_MC_VDAs.Visibility = "Visible"
			    $Enable_Maintenance_MCVDAs.Visibility = "Visible"
			    $Disble_Maintenance_MCVDAs.Visibility = "Visible"
			    $PowerOn_MCVDAs.Visibility = "Visible"
			    $PowerOff_MCVDAs.Visibility = "Visible"
			    $Refresh_MCVDAs.Visibility = "Visible"
			    $Export_MCVDAs.Visibility = "Visible"
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Searching VDAs"
			    Refresh_MCs
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_MC_VDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Enable_Maintenance_MCVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_MC_VDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to enable maintenance for the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "Enable_Maintenance_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_MC_VDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Enable_Maintenance_MCVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Disble_Maintenance_MCVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_MC_VDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $lastButtonClicked = $null
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to disable maintenance for the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "Disble_Maintenance_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_MC_VDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Disble_Maintenance_MCVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOn_MCVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_MC_VDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power on the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "PowerOn_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_MC_VDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOn_MCVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOff_MCVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_MC_VDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power off the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "PowerOff_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_MC_VDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOff_MCVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_MCVDAs.add_Click({
		    try
		    {
			    $datagrid_MC_VDAs.ItemsSource = $null
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing VDAs"
			    Refresh_MCs
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_MCVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_MCVDAs.add_Click({
		    try
		    {
			    if ($datagrid_MC_VDAs.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No VDAs to export." }
			    else
			    {
				    $MC_Name = $datagrid_MCsList.selecteditem.Name
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportVDAs = [hashtable]::Synchronized(@{
						    MC_Name    = $MC_Name
						    ConfigPath = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportVDAs", $SyncHash_ExportVDAs)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $MC_Name = $SyncHash_ExportVDAs.MC_Name
							    $date = get-date -Format MM_dd_yyyy
							    $Export_VDAs = $SyncHash_ExportVDAs.ConfigPath + "\Exports\VDAs_$MC_Name-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_VDAs)
							    {
								    $i++
								    $Export_VDAs = $SyncHash_ExportVDAs.ConfigPath + "\Exports\VDAs_$MC_Name-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_MC_VDAs.ItemsSource | Export-xlsx -Path $Export_VDAs -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_VDAs"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_MCVDAs_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_MCVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_MC_VDA_Simple.add_Click({
		    $Grid_MC_VDAs_Simple.Visibility = "visible"
		    $Grid_MC_VDA_Full.Visibility = "collapse"
	    })
    $Switch_MC_VDAs_Full.add_Click({
		    $Grid_MC_VDAs_Simple.Visibility = "collapse"
		    $Grid_MC_VDA_Full.Visibility = "visible"
	    })
    $MC_Sessions.add_Click({
		    try
		    {
			    $datagrid_MC_sessions.ItemsSource = $null
			    $TextBox_MC_Sessions.text = ""
			    $TextBox_TotalMC_Sessions.text = ""
			    MCs_collapse
			    if ($datagrid_MCsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a Machine Catalog." }
			    else
			    {
				    $Load_TB.Text = "Searching sessions"
				    Refresh_MCSession
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_MC_Sessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Kill_MCSession.add_Click({
		    try
		    {
			    $Sessions = $datagrid_MC_sessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    Else
			    {
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = ($SyncHash.$Session.Farm).DDC
					    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Stop-BrokerSession
				    }
				    Show-Dialog_Main -Foreground "Blue" -Text "Kill command sent.`r`nPlease refresh sessions in few seconds."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Kill_MCSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Hide_MCSession.add_Click({
		    try
		    {
			    $Sessions = $datagrid_MC_sessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    else
			    {
				    $VDI = 0
				    $i = 0
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = $SyncHash.($Session.Farm).DDC
					    $Hidden = $Session.Hidden
					    $Type = $Session.Type
					    if ($Type -match "VDI") { $VDI += 1 }
					    else
					    {
						    if ($Hidden -eq $false)
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$true
						    }
						    else
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$false
						    }
					    }
				    }
				    if ($VDI -ne 0 -and $i -ne 0)
				    {
					    Refresh_MCSession
					    $MainLayer.IsEnabled = $false
					    $Main_MB.FontSize = "20"
					    $Main_MB.text = "Hide status changed for server connections."
					    $Main_MB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = "`r`nBut hide status for VDI can't be changed."; Foreground = "Red" }))
					    $Dialog_Main.IsOpen = $True
					    $Main_MB_Close.add_Click({
							    $Dialog_Main.IsOpen = $False
							    $MainLayer.IsEnabled = $true
						    })
				    }
				    if ($VDI -ne 0 -and $i -eq 0) { Show-Dialog_Main -Foreground "Red" -Text "Hide status for VDI can't be changed." }
				    if ($VDI -eq 0 -and $i -ne 0)
				    {
					    Refresh_MCSession
					    Show-Dialog_Main -Foreground "Blue" -Text "Hide status changed."
				    }
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Hide_MCSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Shadow_MCSession.add_Click({
		    try
		    {
			    $Session = $datagrid_MC_sessions.SelectedItems
			    if ($Session.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    elseif ($Session.count -ge "2") { Show-Dialog_Main -Foreground "Red" -Text "Please select only one session." }
			    else
			    {
				    $Machine = $Session."Machine Name"
				    $User = $Session.User
				    $Domain = $Session.Domain
				    $ID = Get-UserNameSessionIDMap -Comp $Machine | ? { $_.UserName -match $User } | Select-Object -ExpandProperty SessionID
				    $Arg = "/offerra $Machine "
				    $Arg += "$Domain\$User"
				    $Arg += ":"
				    $Arg += $ID
				    Start-Process msra $Arg
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Shadow_MCSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_MCSession.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing sessions"
			    Refresh_MCSession
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_MCSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_MCSession.add_Click({
		    try
		    {
			    if ($datagrid_MC_sessions.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No sessions to export." }
			    else
			    {
				    $MC_Name = $datagrid_MCsList.selecteditem.Name
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportMCSessions = [hashtable]::Synchronized(@{
						    MC_Name    = $MC_Name
						    ConfigPath = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportMCSessions", $SyncHash_ExportMCSessions)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $Name = $SyncHash_ExportMCSessions.MC_Name
							    $date = get-date -Format MM_dd_yyyy
							    $Export_MCSessions = $SyncHash_ExportMCSessions.ConfigPath + "\Exports\Sessions_$Name-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_MCSessions)
							    {
								    $i++
								    $Export_MCSessions = $SyncHash_ExportMCSessions.ConfigPath + "\Exports\Sessions_$Name-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_MC_sessions.ItemsSource | Export-xlsx -Path $Export_MCSessions -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_MCSessions"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_MCSession_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_MCSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $MC_Refresh.add_Click({
		    if ($S_MCs_Details.selectedItem -ne $null)
		    {
			    $Load_TB.Text = "Refreshing Machine Catalogs"
			    S_MCs_Details
		    }
		    else
		    {
			    Search_MC
			    $Load_TB.Text = "Refreshing Machine Catalogs"
		    }
	    })
    $Switch_MC_Session_Simple.add_Click({
		    $Grid_MC_Sessions_Simple.Visibility = "visible"
		    $Grid_MC_Sessions_Full.Visibility = "collapse"
	    })
    $Switch_MC_Sessions_Full.add_Click({
		    $Grid_MC_Sessions_Simple.Visibility = "collapse"
		    $Grid_MC_Sessions_Full.Visibility = "visible"
	    })
    $S_MCs_Details.Add_SelectionChanged({
		    $Load_TB.Text = "Searching Machine Catalogs"
		    S_MCs_Details
	    })
    #####################
    # End_MachineCatalogs
    #####################
    ######################
    # Start_DeliveryGroups
    ######################
    $DG_TB.Add_KeyDown({
		    param ($sender,
			    $e)
		    if ($e.Key -eq [System.Windows.Input.Key]::Enter)
		    {
			    $Load_TB.Text = "Searching Delivery Groups"
			    Search_DG
		    }
	    })
    $Search_DG_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for a Delivery Group.") })
    $Search_AllDGs_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for all Delivery Group in the farm selected.") })
    $Search_DG.add_Click({
		    $Load_TB.Text = "Searching Delivery Groups"
		    Search_DG
	    })
    $DG_Details.add_Click({
		    try
		    {
			    DGs_collapse
			    if ($datagrid_DGsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a Delivery Group." }
			    else
			    {
				    $datagrid_DG_settings.ItemsSource = $null
				    $ListView_DG_Desktops.ItemsSource = $null
				    $ListView_DG_Desktops.Items.Clear()
				    $ListView_DG_Reboot.ItemsSource = $null
				    $ListView_DG_Reboot.Items.Clear()
				    $Farm = $datagrid_DGsList.selecteditem.farm
				    $UID = $datagrid_DGsList.selecteditem.uid
				    $Name = $datagrid_DGsList.selecteditem.name
				    $DDC = ($SyncHash.$Farm).DDC
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Searching Delivery Group Details"
				    $Global:SyncHash_DG_Details = [hashtable]::Synchronized(@{
						    Farm				 = $Farm
						    UID				     = $UID
						    DDC				     = $DDC
						    Name				 = $Name
						    datagrid_DG_settings = $datagrid_DG_settings
						    Grid_DG			     = $Grid_DG
						    ListView_DG_Desktops = $ListView_DG_Desktops
						    ListView_DG_Reboot   = $ListView_DG_Reboot
						    DG_PublishedDesktops = $DG_PublishedDesktops
						    DG_RebootSchedule    = $DG_RebootSchedule
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_DG_Details", $SyncHash_DG_Details)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    asnp Citrix*
							    $DG = Get-BrokerDesktopGroup -AdminAddress $SyncHash_DG_Details.DDC -Uid $SyncHash_DG_Details.UID
							    $DG_VDAs = $DG.TotalDesktops
							    $DG_Name = $SyncHash_DG_Details.Name
							    $DG_Description = $DG.Description
							    $DG_PublishedName = $DG.PublishedName
							    $DG_Farm = $SyncHash_DG_Details.Farm
							    $DG_SessionSupport = $DG.SessionSupport
							    $DG_DesktopKind = $DG.DesktopKind
							    $DG_TotalApplications = $DG.TotalApplications
							    $global:SyncHash_DG_Details.DG_PublishedDesktops = Get-BrokerEntitlementPolicyRule -AdminAddress $SyncHash_DG_Details.DDC -DesktopGroupUid $SyncHash_DG_Details.UID | Select-Object @{ n = "Name"; e = { $_.Name } }, @{ n = "UID"; e = { $_.UID } }
							    $DG_DeliveryType = $DG.DeliveryType
							    $Global:SyncHash_DG_Details.DG_RebootSchedule = Get-BrokerRebootScheduleV2 -AdminAddress $SyncHash_DG_Details.DDC -DesktopGroupUid $SyncHash_DG_Details.UID | Select-Object @{ n = "Name"; e = { $_.Name } }, @{ n = "UID"; e = { $_.UID } }
							    $SyncHash_DG_Details.datagrid_DG_settings.Dispatcher.Invoke([Action]{ $SyncHash_DG_Details.datagrid_DG_settings.Visibility = "Visible" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash_DG_Details.Grid_DG.Dispatcher.Invoke([Action]{ $SyncHash_DG_Details.Grid_DG.Visibility = "Visible" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash_DG_Details.datagrid_DG_settings.Dispatcher.Invoke([Action]{
									    $SyncHash_DG_Details.datagrid_DG_settings.ItemsSource = @(
										    [PSCustomObject]@{ Column1Header = "Published Name"; Column2Data = $DG_PublishedName; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Description"; Column2Data = $DG_Description; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Session Support"; Column2Data = $DG_SessionSupport; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Desktop Kind"; Column2Data = $DG_DesktopKind; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Delivery Type"; Column2Data = $DG_DeliveryType; IsReadOnly = $true },
										    [PSCustomObject]@{ Column1Header = "Total Applications"; Column2Data = $DG_TotalApplications; IsReadOnly = $true })
								    }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Details_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $SyncHash_DG_Details.DG_PublishedDesktops = Get-BrokerEntitlementPolicyRule -AdminAddress $SyncHash_DG_Details.DDC -DesktopGroupUid $SyncHash_DG_Details.UID | Select-Object @{ n = "Name"; e = { $_.Name } }, @{ n = "UID"; e = { $_.UID } }
				    $SyncHash_DG_Details.DG_RebootSchedule = Get-BrokerRebootScheduleV2 -AdminAddress $SyncHash_DG_Details.DDC -DesktopGroupUid $SyncHash_DG_Details.UID | Select-Object @{ n = "Name"; e = { $_.Name } }, @{ n = "UID"; e = { $_.UID } }
				
				    foreach ($item in $Global:SyncHash_DG_Details.DG_PublishedDesktops)
				    {
					    $listView_DG_DesktopItem = New-Object System.Windows.Controls.ListViewItem
					    $listView_DG_DesktopItem.Content = $item
					    $ListView_DG_Desktops.Items.Add($listView_DG_DesktopItem)
				    }
				    foreach ($item in $Global:SyncHash_DG_Details.DG_RebootSchedule)
				    {
					    $listView_DG_RebootScheduleItem = New-Object System.Windows.Controls.ListViewItem
					    $listView_DG_RebootScheduleItem.Content = $item
					    $ListView_DG_Reboot.Items.Add($listView_DG_RebootScheduleItem)
				    }
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Details " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Desktop_Settings.add_Click({
		    try
		    {
			    if ($ListView_DG_Desktops.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a published desktop." }
			    else
			    {
				    $Grid_DG_Reboot_settings.Visibility = "Collapse"
				    $Farm = $datagrid_DGsList.selecteditem.farm
				    $DDC = ($SyncHash.$Farm).DDC
				    $UID = $ListView_DG_Desktops.selecteditem.Content.UID
				    $App = Get-BrokerEntitlementPolicyRule -AdminAddress $DDC -Uid $UID
				    $App_PublishedName = $App.PublishedName
				    $App_BrowserName = $App.BrowserName
				    $App_Name = $App.Name
				    $App_Description = $App.Description
				    $App_Enabled = $App.Enabled
				    $App_RestrictToTag = $App.RestrictToTag
				    $Grid_DG_Desk_settings.Visibility = "Visible"
				    $datagrid_DG_Desk_settings.ItemsSource = @(
					    [PSCustomObject]@{ Column1Header = "Browser name"; Column2Data = $App_BrowserName; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Name"; Column2Data = $App_Name; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "PublishedName / Display name"; Column2Data = $App_PublishedName; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Description"; Column2Data = $App_Description; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Restrict To Tag"; Column2Data = $App_RestrictToTag; IsReadOnly = $True }
				    )
				    $datagrid_DG_Desk_settings_2.ItemsSource = @([PSCustomObject]@{ Column1Header = "Enabled"; Column2Data = [System.Collections.ObjectModel.ObservableCollection[object]]@($true, $false); Column2SelectedValue = $App_Enabled; IsReadOnly = $false })
				    $TAG_List = @()
				    $TAG_List = Get-BrokerTag -AdminAddress $DDC | Select-Object -ExpandProperty Name
				    foreach ($tag in $TAG_List) { $listbox_DG_Desk_tag.Items.Add($tag) }
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Desktop_Settings_DG " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $DG_Desk_settings_Apply.add_Click({
		    try
		    {
			    $Farm = $datagrid_DGsList.selecteditem.Farm
			    $DDC = ($SyncHash.$Farm).DDC
			    $UID = $ListView_DG_Desktops.selecteditem.Content.UID
			    Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -PublishedName $datagrid_DG_Desk_settings.itemssource.Column2Data[2]
			    Rename-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -NewName $datagrid_DG_Desk_settings.itemssource.Column2Data[1]
			    Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -Description $datagrid_DG_Desk_settings.itemssource.Column2Data[3]
			    if ($datagrid_DG_Desk_settings_2.itemssource.Column2SelectedValue[0] -eq $True) { Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -Enabled $True }
			    else { Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -Enabled $False }
			    Show-Dialog_Main -Foreground "Blue" -Text "Changes applied"
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Desk_settings_Apply " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $DG_Desk_settings_Discard.add_Click({
		    try
		    {
			    $datagrid_DG_Desk_settings.ItemsSource = $null
			    $datagrid_DG_Desk_settings_2.ItemsSource = $null
			    $listbox_DG_Desk_tag.ItemsSource = $null
			    $listbox_DG_Desk_tag.Items.Clear()
			    $Farm = $datagrid_DGsList.selecteditem.farm
			    $DDC = ($SyncHash.$Farm).DDC
			    $UID = $ListView_DG_Desktops.selecteditem.Content.UID
			    $App = Get-BrokerEntitlementPolicyRule -AdminAddress $DDC -Uid $UID
			    $App_PublishedName = $App.PublishedName
			    $App_BrowserName = $App.BrowserName
			    $App_Name = $App.Name
			    $App_Description = $App.Description
			    $App_Enabled = $App.Enabled
			    $App_RestrictToTag = $App.RestrictToTag
			    $Grid_DG_Desk_settings.Visibility = "Visible"
			    $datagrid_DG_Desk_settings.ItemsSource = @(
				    [PSCustomObject]@{ Column1Header = "Browser name"; Column2Data = $App_BrowserName; IsReadOnly = $true },
				    [PSCustomObject]@{ Column1Header = "Name"; Column2Data = $App_Name; IsReadOnly = $false },
				    [PSCustomObject]@{ Column1Header = "PublishedName / Display name"; Column2Data = $App_PublishedName; IsReadOnly = $false },
				    [PSCustomObject]@{ Column1Header = "Description"; Column2Data = $App_Description; IsReadOnly = $false },
				    [PSCustomObject]@{ Column1Header = "Restrict To Tag"; Column2Data = $App_RestrictToTag; IsReadOnly = $True }
			    )
			    $datagrid_DG_Desk_settings_2.ItemsSource = @([PSCustomObject]@{ Column1Header = "Enabled"; Column2Data = [System.Collections.ObjectModel.ObservableCollection[object]]@($true, $false); Column2SelectedValue = $App_Enabled; IsReadOnly = $false })
			    $TAG_List = @()
			    $TAG_List = Get-BrokerTag -AdminAddress $DDC | Select-Object -ExpandProperty Name
			    foreach ($tag in $TAG_List) { $listbox_DG_Desk_tag.Items.Add($tag) }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Desk_settings_Discard " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $DG_Desk_tag.add_Click({
		    try
		    {
			    $Farm = $datagrid_DGsList.selecteditem.Farm
			    $DDC = ($SyncHash.$Farm).DDC
			    $UID = $ListView_DG_Desktops.selecteditem.Content.UID
			    If ($listbox_DG_Desk_tag.SelectedItem -eq $Null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a TAG to add or change." }
			    else
			    {
				    Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -RestrictToTag $listbox_DG_Desk_tag.SelectedItem
				    Show-Dialog_Main -Foreground "Blue" -Text "TAG modified.`r`nPlease refresh."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Desk_tag " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $DG_Desk_tag_remove.add_Click({
		    try
		    {
			    $Farm = $datagrid_DGsList.selecteditem.Farm
			    $DDC = ($SyncHash.$Farm).DDC
			    $UID = $ListView_DG_Desktops.selecteditem.Content.UID
			    if ($datagrid_DG_Desk_settings.itemssource.Column2Data[4] -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "No TAG to remove." }
			    else
			    {
				    Set-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID -RestrictToTag $Null
				    Show-Dialog_Main -Foreground "Blue" -Text "TAG removed.`r`nPlease refresh."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Desk_tag_remove " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Reboot_Setings.add_Click({
		    try
		    {
			    if ($ListView_DG_Reboot.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a scheduled reboot." }
			    else
			    {
				    $Grid_DG_Desk_settings.Visibility = "Collapse"
				    $Farm = $datagrid_DGsList.selecteditem.farm
				    $DDC = ($SyncHash.$Farm).DDC
				    $UID = $ListView_DG_Reboot.selecteditem.Content.UID
				    $Reboot = Get-BrokerRebootScheduleV2 -AdminAddress $DDC -UID $UID
				    $Reboot_Name = $Reboot.Name
				    $Reboot_Enabled = $Reboot.Enabled
				    $Reboot_Description = $Reboot.Description
				    $Reboot_Frequency = $Reboot.Frequency
				    $Reboot_Day = $Reboot.Day
				    $Reboot_StartTime = $Reboot.StartTime
				    $Reboot_RebootDuration = $Reboot.RebootDuration
				    $Reboot_WarningDuration = $Reboot.WarningDuration
				    $Reboot_WarningMessage = $Reboot.WarningMessage
				    $Reboot_WarningRepeatInterval = $Reboot.WarningRepeatInterval
				    $Reboot_WarningTitle = $Reboot.WarningTitle
				    $Grid_DG_Reboot_settings.Visibility = "Visible"
				    $datagrid_DG_Reboot_settings.ItemsSource = @(
					    [PSCustomObject]@{ Column1Header = "Name"; Column2Data = $Reboot_Name; IsReadOnly = $true },
					    [PSCustomObject]@{ Column1Header = "Enabled"; Column2Data = $Reboot_Enabled; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Description / Display name"; Column2Data = $Reboot_Description; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Frequency"; Column2Data = $Reboot_Frequency; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Day"; Column2Data = $Reboot_Day; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Start Time"; Column2Data = $Reboot_StartTime; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Reboot Duration"; Column2Data = $Reboot_RebootDuration; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Warning Duration"; Column2Data = $Reboot_WarningDuration; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Warning Message"; Column2Data = $Reboot_WarningMessage; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Warning Repeat Interval"; Column2Data = $Reboot_WarningMessage; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Warning Title"; Column2Data = $Reboot_WarningTitle; IsReadOnly = $false },
					    [PSCustomObject]@{ Column1Header = "Frequency"; Column2Data = $Reboot_Frequency; IsReadOnly = $false }
				    )
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Reboot_Setings " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $DG_VDAs.add_Click({
		    try
		    {
			    $datagrid_DG_VDAs.ItemsSource = $null
			    $TextBox_DG_VDAs.text = ""
			    $TextBox_TotalDG_VDAs.text = ""
			    DGs_collapse
			    if ($datagrid_DGsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a Delivery Group." }
			    else
			    {
				    $datagrid_DG_VDAs.Visibility = "Visible"
				    $Enable_Maintenance_DGVDAs.Visibility = "Visible"
				    $Disble_Maintenance_DGVDAs.Visibility = "Visible"
				    $PowerOn_DGVDAs.Visibility = "Visible"
				    $PowerOff_DGVDAs.Visibility = "Visible"
				    $Refresh_DGVDAs.Visibility = "Visible"
				    $Export_DGVDAs.Visibility = "Visible"
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Searching VDAs"
				    Refresh_DGs
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_VDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Enable_Maintenance_DGVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_DG_VDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to enable maintenance for the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "Enable_Maintenance_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_DG_VDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Enable_Maintenance_DGVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Disble_Maintenance_DGVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_DG_VDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $lastButtonClicked = $null
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to disable maintenance for the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "Disble_Maintenance_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_DG_VDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Disble_Maintenance_DGVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOn_DGVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_DG_VDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power on the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "PowerOn_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_DG_VDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOn_DGVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOff_DGVDAs.add_Click({
		    try
		    {
			    $VDAs = $datagrid_DG_VDAs.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power off the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "PowerOff_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_DG_VDAs.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOff_DGVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_DGVDAs.add_Click({
		    try
		    {
			    $datagrid_DG_VDAs.ItemsSource = $null
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing VDAs"
			    Refresh_DGs
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_DGVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_DGVDAs.add_Click({
		    try
		    {
			    if ($datagrid_DG_VDAs.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No VDAs to export." }
			    else
			    {
				    $DG_Name = $datagrid_DGsList.selecteditem.Name
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportVDAs = [hashtable]::Synchronized(@{
						    DG_Name    = $DG_Name
						    ConfigPath = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportVDAs", $SyncHash_ExportVDAs)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $DG_Name = $SyncHash_ExportVDAs.DG_Name
							    $date = get-date -Format MM_dd_yyyy
							    $Export_VDAs = $SyncHash_ExportVDAs.ConfigPath + "\Exports\VDAs_$DG_Name-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_VDAs)
							    {
								    $i++
								    $Export_VDAs = $SyncHash_ExportVDAs.ConfigPath + "\Exports\VDAs_$DG_Name-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_DG_VDAs.ItemsSource | Export-xlsx -Path $Export_VDAs -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_VDAs"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_DGVDAs_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_DGVDAs " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_DG_VDA_Simple.add_Click({
		    $Grid_DG_VDAs_Simple.Visibility = "visible"
		    $Grid_DG_VDA_Full.Visibility = "collapse"
	    })
    $Switch_DG_VDAs_Full.add_Click({
		    $Grid_DG_VDAs_Simple.Visibility = "collapse"
		    $Grid_DG_VDA_Full.Visibility = "visible"
	    })
    $DG_Sessions.add_Click({
		    try
		    {
			    $datagrid_DG_sessions.ItemsSource = $null
			    $TextBox_DG_Sessions.text = ""
			    $TextBox_TotalDG_Sessions.text = ""
			    DGs_collapse
			    if ($datagrid_DGsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a Delivery Group." }
			    else
			    {
				    $Load_TB.Text = "Searching sessions"
				    Refresh_DGSession
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Sessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Kill_DGSession.add_Click({
		    try
		    {
			    $Sessions = $datagrid_DG_sessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    Else
			    {
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = ($SyncHash.$Session.Farm).DDC
					    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Stop-BrokerSession
				    }
				    Show-Dialog_Main -Foreground "Blue" -Text "Kill command sent.`r`nPlease refresh sessions in few seconds."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Kill_DGSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Hide_DGSession.add_Click({
		    try
		    {
			    $Sessions = $datagrid_DG_sessions.SelectedItems
			    if ($Sessions.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    else
			    {
				    $VDI = 0
				    $i = 0
				    foreach ($Session in $Sessions)
				    {
					    $UID = $Session.Uid
					    $DDC = $SyncHash.($Session.Farm).DDC
					    $Hidden = $Session.Hidden
					    $Type = $Session.Type
					    if ($Type -match "VDI") { $VDI += 1 }
					    else
					    {
						    if ($Hidden -eq $false)
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$true
						    }
						    else
						    {
							    $i += 1
							    Get-BrokerSession -AdminAddress $DDC -UiD $UID | Set-BrokerSession -Hidden:$false
						    }
					    }
				    }
				    if ($VDI -ne 0 -and $i -ne 0)
				    {
					    Refresh_DGSession
					    $MainLayer.IsEnabled = $false
					    $Main_MB.FontSize = "20"
					    $Main_MB.text = "Hide status changed for server connections."
					    $Main_MB.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{ Text = "`r`nBut hide status for VDI can't be changed."; Foreground = "Red" }))
					    $Dialog_Main.IsOpen = $True
					    $Main_MB_Close.add_Click({
							    $Dialog_Main.IsOpen = $False
							    $MainLayer.IsEnabled = $true
						    })
				    }
				    if ($VDI -ne 0 -and $i -eq 0) { Show-Dialog_Main -Foreground "Red" -Text "Hide status for VDI can't be changed." }
				    if ($VDI -eq 0 -and $i -ne 0)
				    {
					    Refresh_DGSession
					    Show-Dialog_Main -Foreground "Blue" -Text "Hide status changed."
				    }
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Hide_DGSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Shadow_DGSession.add_Click({
		    try
		    {
			    $Session = $datagrid_DG_sessions.SelectedItems
			    if ($Session.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a session." }
			    elseif ($Session.count -ge "2") { Show-Dialog_Main -Foreground "Red" -Text "Please select only one session." }
			    else
			    {
				    $Machine = $Session."Machine Name"
				    $User = $Session.User
				    $Domain = $Session.Domain
				    $ID = Get-UserNameSessionIDMap -Comp $Machine | ? { $_.UserName -match $User } | Select-Object -ExpandProperty SessionID
				    $Arg = "/offerra $Machine "
				    $Arg += "$Domain\$User"
				    $Arg += ":"
				    $Arg += $ID
				    Start-Process msra $Arg
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Shadow_DGSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_DGSession.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing sessions"
			    Refresh_DGSession
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Sessions " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_DGSession.add_Click({
		    try
		    {
			    if ($datagrid_DG_sessions.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No sessions to export." }
			    else
			    {
				    $DG_Name = $datagrid_DGsList.selecteditem.Name
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportDGSessions = [hashtable]::Synchronized(@{
						    DG_Name    = $DG_Name
						    ConfigPath = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportDGSessions", $SyncHash_ExportDGSessions)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $Name = $SyncHash_ExportDGSessions.DG_Name
							    $date = get-date -Format MM_dd_yyyy
							    $Export_DGSessions = $SyncHash_ExportDGSessions.ConfigPath + "\Exports\Sessions_$Name-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_DGSessions)
							    {
								    $i++
								    $Export_DGSessions = $SyncHash_ExportDGSessions.ConfigPath + "\Exports\Sessions_$Name-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_DG_sessions.ItemsSource | Export-xlsx -Path $Export_DGSessions -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_DGSessions"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_DGSession_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_DGSession " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_DG_Session_Simple.add_Click({
		    $Grid_DG_Sessions_Simple.Visibility = "visible"
		    $Grid_DG_Sessions_Full.Visibility = "collapse"
	    })
    $Switch_DG_Sessions_Full.add_Click({
		    $Grid_DG_Sessions_Simple.Visibility = "collapse"
		    $Grid_DG_Sessions_Full.Visibility = "visible"
	    })
    $DG_Publications.add_Click({
		    try
		    {
			    $datagrid_DG_Publications.ItemsSource = $null
			    $TextBox_DG_Publications.text = ""
			    $TextBox_TotalDG_Publications.text = ""
			    DGs_collapse
			    if ($datagrid_DGsList.selecteditem -eq $null) { Show-Dialog_Main -Foreground "Red" -Text "Please select a Delivery Group." }
			    else
			    {
				    $Load_TB.Text = "Searching publications"
				    Refresh_DGPublication
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Publications " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $DG_Refresh.add_Click({
		    if ($S_DGs_Details.selectedItem -ne $null)
		    {
			    $Load_TB.Text = "Refreshing Delivery Groups"
			    S_DGs_Details
		    }
		    else
		    {
			    Search_DG
			    $Load_TB.Text = "Refreshing Delivery Groups"
		    }
	    })
    $Disable_DGPublis.add_Click({
		    try
		    {
			    $Publis = $datagrid_DG_Publications.SelectedItems
			    if ($Publis.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    Else
			    {
				    foreach ($Publi in $Publis)
				    {
					    $UID = $Publi.Uid
					    $Enabled = $Publi.Enabled
					    $DDC = ($SyncHash.$Publi.Farm).DDC
					    if ($Enabled -eq $false) { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Enabled $true }
					    else { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Enabled $false }
				    }
				    $Load_TB.Text = "Refreshing publications"
				    Refresh_DGPublication
				    Show-Dialog_Main -Foreground "Blue" -Text "Enabled status changed for the selected publications."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Disable_DGPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Hide_DGPublis.add_Click({
		    try
		    {
			    $Publis = $datagrid_DG_Publications.SelectedItems
			    if ($Publis.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    Else
			    {
				    foreach ($Publi in $Publis)
				    {
					    $UID = $Publi.Uid
					    $Visible = $Publi.Visible
					    $DDC = ($SyncHash.$Publi.Farm).DDC
					    if ($Visible -eq $false) { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Visible $true }
					    else { Set-BrokerApplication -AdminAddress $DDC -InputObject $UID -Visible $false }
				    }
				    $Load_TB.Text = "Refreshing publications"
				    Refresh_DGPublication
				    Show-Dialog_Main -Foreground "Blue" -Text "Visible state changed for the selected publications."
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Hide_DGPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
		
	    })
    $Delete_DGPublis.add_Click({
		    try
		    {
			    $Publis = $datagrid_DG_Publications.SelectedItems
			    if ($Publis.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a publication." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to delete the selected publications ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $Main_MB_Confirm.add_Click({
						    $Publis = $datagrid_AllPublications.SelectedItems
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
						    foreach ($Publi in $Publis)
						    {
							    $UID = $Publi.Uid
							    $Type = $Publi.Type
							    $DDC = ($SyncHash.$Publi.Farm).DDC
							    if ($Type -eq "Application") { Remove-BrokerApplication -AdminAddress $DDC -InputObject $UID }
							    else { Remove-BrokerEntitlementPolicyRule -AdminAddress $DDC -InputObject $UID }
						    }
						    $Load_TB.Text = "Refreshing publications"
						    Refresh_DGPublication
						    Show-Dialog_Main -Foreground "Blue" -Text "Selected publications deleted."
					    })
				    $Main_MB_Cancel.add_Click({
						    $Dialog_Main_Confirm.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Delete_DGPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_DGPublis.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing publications"
			    Refresh_DGPublication
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_DG_Publications " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_DGPublis.add_Click({
		    try
		    {
			    if ($datagrid_DG_Publications.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No publications to export." }
			    else
			    {
				    $DGSelected = $datagrid_DGsList.selectedItem.name
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportPublications = [hashtable]::Synchronized(@{
						    DGSelected = $DGSelected
						    ConfigPath = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportPublications", $SyncHash_ExportPublications)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $DG = $SyncHash_ExportPublications.DGSelected
							    $date = get-date -Format MM_dd_yyyy
							    $Export_Publications = $SyncHash_ExportPublications.ConfigPath + "\Exports\Publications_$DG-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_Publications)
							    {
								    $i++
								    $Export_Publications = $SyncHash_ExportPublications.ConfigPath + "\Exports\Publications_$DG-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_AllPublications.ItemsSource | Export-xlsx -Path $Export_Publications -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_Publications"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_DGPublis_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_DGPublis " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_DG_Publications_Simple.add_Click({
		    $Grid_DG_Publications_Simple.Visibility = "visible"
		    $Grid_DG_Publications_Full.Visibility = "collapse"
	    })
    $Switch_DG_Publications_Full.add_Click({
		    $Grid_DG_Publications_Simple.Visibility = "collapse"
		    $Grid_DG_Publications_Full.Visibility = "visible"
	    })
    $S_DGs_Details.Add_SelectionChanged({
		    $Load_TB.Text = "Searching Delivery Groups"
		    S_DGs_Details
	    })
    ####################
    # End_DeliveryGroups
    ####################
    ################################
    # Start_Maintenance_Registration
    ################################
    $Maintenance_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for VDAs in maintenance mode.") })
    $Registration_Help.add_Click({ $snackbar.MessageQueue.Enqueue("Use this part to search for unregistered VDAs.") })
    $S_Maintenance.Add_SelectionChanged({
		    try
		    {
			    if ($S_Maintenance.selectedItem -ne $null)
			    {
				    Refresh_Maintenance
				    $S_Registration.SelectedItem = $null
				    $Load_TB.Text = "Searching VDAs in maintenance mode"
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_Maintenance " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $S_Registration.Add_SelectionChanged({
		    try
		    {
			    if ($S_Registration.selectedItem -ne $null)
			    {
				    Refresh_Registration
				    $S_Maintenance.SelectedItem = $null
				    $Load_TB.Text = "Searching unregistered VDAs"
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_Registration " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_MaintRegist_Simple.add_Click({
		    $Grid_Simple_Maintenance_Registered.Visibility = "visible"
		    $Grid_Detailled_Maintenance_Registered.Visibility = "collapse"
	    })
    $Refresh_Maintenance_Simple.add_Click({
		    try
		    {
			    $Grid_Simple_Maintenance_Registered.Visibility = "collapse"
			    $Grid_Detailled_Maintenance_Registered.Visibility = "collapse"
			    $Refresh_Maintenance.Visibility = "collapse"
			    $Refresh_Registration.Visibility = "collapse"
			    $Refresh_Maintenance_Simple.Visibility = "collapse"
			    $Refresh_Registration_Simple.Visibility = "collapse"
			    $datagrid_Maintenance_Registered.ItemsSource = $null
			    $TextBox_Servers_Maintenance_Registered.Text = ""
			    $TextBox_TotalServers_Maintenance_Registered.Text = ""
			    $Farm = $S_Maintenance.selectedItem
			    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing VDAs in maintenance mode"
			    $Global:SyncHash_Maint_list = [hashtable]::Synchronized(@{
					    Farm	  = $Farm
					    DDC	      = $DDC
					    Farm_List = $SyncHash.Farm_List
				    })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_Maint_list", $SyncHash_Maint_list)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $Maint_list = @()
						    $Total_Maint = @()
						    if ($SyncHash_Maint_list.Farm -eq "All Farms")
						    {
							    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
							    foreach ($item in $SyncHash.Farm_List)
							    {
								    $DDC = ($SyncHash.$item).DDC
								    $Maint_list += Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $DDC | Where-Object { $_.InMaintenanceMode -eq $true } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $item } }, @{
									    n = "Type"; e = {
										    if ($_.SessionSupport -match "MultiSession") { "Server" }
										    else { "VDI" }
									    }
								    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
							    }
							    $Maint_list = $Maint_list | Sort-Object "Machine Name"
							    $Total_Maint = $Maint_list."Machine Name".count
							    $Simple_Maint_List = $Maint_list."Machine Name"
							    if ($Total_Maint -eq 0) { $Simple_Maint_List_String = $null }
							    else { $Simple_Maint_List_String = [string]::Join([Environment]::NewLine, $Simple_Maint_List) }
						    }
						    else
						    {
							    $Maint_list = Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $SyncHash_Maint_list.DDC | Where-Object { $_.InMaintenanceMode -eq $true } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $SyncHash_Maint_list.Farm } }, @{
								    n = "Type"; e = {
									    if ($_.SessionSupport -match "MultiSession") { "Server" }
									    else { "VDI" }
								    }
							    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
							    $Maint_list = $Maint_list | Sort-Object "Machine Name"
							    $Total_Maint = $Maint_list."Machine Name".count
							    $Simple_Maint_List = $Maint_list."Machine Name"
							    if ($Total_Maint -eq 0) { $Simple_Maint_List_String = $null }
							    else { $Simple_Maint_List_String = [string]::Join([Environment]::NewLine, $Simple_Maint_List) }
						    }
						    if ($Total_Maint -eq 0)
						    {
							    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Maint" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Grid_Simple_Maintenance_Registered.Visibility = "Visible"
									    $SyncHash.Refresh_Maintenance.Visibility = "Visible"
									    $SyncHash.Refresh_Maintenance_Simple.Visibility = "Visible"
									    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $false
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No VDA found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($Total_Maint -eq 1)
						    {
							    $Maint_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $Maint_List_Datagrid.Add($Maint_list)
							    $SyncHash.datagrid_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_Maintenance_Registered.ItemsSource = $Maint_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.TextBox_Servers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Maintenance_Registered.text = $Simple_Maint_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Maint" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Grid_Simple_Maintenance_Registered.Visibility = "Visible"
									    $SyncHash.Refresh_Maintenance.Visibility = "Visible"
									    $SyncHash.Refresh_Maintenance_Simple.Visibility = "Visible"
									    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $false
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
						    else
						    {
							    $Maint_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $Maint_List_Datagrid.AddRange($Maint_list)
							    $SyncHash.datagrid_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_Maintenance_Registered.ItemsSource = $Maint_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.TextBox_Servers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Maintenance_Registered.text = $Simple_Maint_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Maint" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Grid_Simple_Maintenance_Registered.Visibility = "Visible"
									    $SyncHash.Refresh_Maintenance.Visibility = "Visible"
									    $SyncHash.Refresh_Maintenance_Simple.Visibility = "Visible"
									    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $false
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Maintenance_Simple_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Maintenance_Simple " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_Registration_Simple.add_Click({
		    try
		    {
			    $Grid_Simple_Maintenance_Registered.Visibility = "collapse"
			    $Grid_Detailled_Maintenance_Registered.Visibility = "collapse"
			    $Refresh_Maintenance.Visibility = "collapse"
			    $Refresh_Registration.Visibility = "collapse"
			    $Refresh_Maintenance_Simple.Visibility = "collapse"
			    $Refresh_Registration_Simple.Visibility = "collapse"
			    $datagrid_Maintenance_Registered.ItemsSource = $null
			    $TextBox_Servers_Maintenance_Registered.Text = ""
			    $TextBox_TotalServers_Maintenance_Registered.Text = ""
			    $Farm = $S_Registration.selectedItem
			    if ($Farm -ne $null) { $DDC = ($SyncHash.$Farm).DDC }
			    $MainLayer.IsEnabled = $false
			    $SpinnerOverlayLayer_Main.Visibility = "Visible"
			    $Load_TB.Text = "Refreshing unregistered VDAs"
			    $Global:SyncHash_Regist_list = [hashtable]::Synchronized(@{
					    Farm	  = $Farm
					    DDC	      = $DDC
					    Farm_List = $SyncHash.Farm_List
				    })
			    $Runspace = [runspacefactory]::CreateRunspace()
			    $Runspace.ThreadOptions = "ReuseThread"
			    $Runspace.ApartmentState = "STA"
			    $Runspace.Open()
			    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
			    $Runspace.SessionStateProxy.SetVariable("SyncHash_Regist_list", $SyncHash_Regist_list)
			    $Worker = [PowerShell]::Create().AddScript({
					    try
					    {
						    asnp Citrix*
						    $Maint_list = @()
						    $Total_Maint = @()
						    if ($SyncHash_Regist_list.Farm -eq "All Farms")
						    {
							    $SyncHash.Farm_List = $SyncHash.Farm_List | Where-Object { $_ -ne "All farms" }
							    foreach ($item in $SyncHash.Farm_List)
							    {
								    $DDC = ($SyncHash.$item).DDC
								    $Regist_list += Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $DDC | Where-Object { $_.RegistrationState -eq "Unregistered" } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $item } }, @{
									    n = "Type"; e = {
										    if ($_.SessionSupport -match "MultiSession") { "Server" }
										    else { "VDI" }
									    }
								    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
							    }
							    $Regist_list = $Regist_list | Sort-Object "Machine Name"
							    $Total_Regist = $Regist_list."Machine Name".count
							    $Simple_Regist_List = $Regist_list."Machine Name"
							    if ($Total_Regist -eq 0) { $Simple_Regist_List_String = $null }
							    else { $Simple_Regist_List_String = [string]::Join([Environment]::NewLine, $Simple_Regist_List) }
						    }
						    else
						    {
							    $Regist_list = Get-BrokerMachine -MaxRecordCount 999999 -AdminAddress $SyncHash_Regist_list.DDC | Where-Object { $_.RegistrationState -eq "Unregistered" } | Select-Object @{ n = "Machine Name"; e = { $_.MachineName.Split('\')[-1] } }, @{ n = "Domain"; e = { $_.MachineName.Split('\')[0] } }, @{ n = "Farm"; e = { $SyncHash_Regist_list.Farm } }, @{
								    n = "Type"; e = {
									    if ($_.SessionSupport -match "MultiSession") { "Server" }
									    else { "VDI" }
								    }
							    }, @{ n = "Registration State"; e = { $_.RegistrationState } }, @{ n = "Maintenance State"; e = { $_.InMaintenanceMode } }, @{ n = "Power State"; e = { $_.PowerState } }, @{ n = "Delivery Group"; e = { $_.DesktopGroupName } }, @{ n = "Machine Catalog"; e = { $_.CatalogName } }, @{ n = "Sessions"; e = { $_.SessionCount } }, @{ n = "OS Type"; e = { $_.OSType } }, @{ n = "IP Address"; e = { $_.IPAddress } }, @{ n = "Load"; e = { $_.LoadIndex } }, @{ n = "Tags"; e = { $_.Tags -join ', ' } }, @{ n = "Agent Version"; e = { $_.AgentVersion } }, @{ n = "Provisioning Type"; e = { $_.ProvisioningType } }, UID
							    $Regist_list = $Regist_list | Sort-Object "Machine Name"
							    $Total_Regist = $Regist_list."Machine Name".count
							    $Simple_Regist_List = $Regist_list."Machine Name"
							    if ($Total_Regist -eq 0) { $Simple_Regist_List_String = $null }
							    else { $Simple_Regist_List_String = [string]::Join([Environment]::NewLine, $Simple_Regist_List) }
						    }
						    if ($Total_Regist -eq 0)
						    {
							    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Regist" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Grid_Simple_Maintenance_Registered.Visibility = "Visible"
									    $SyncHash.Refresh_Registration.Visibility = "Visible"
									    $SyncHash.Refresh_Registration_Simple.Visibility = "Visible"
									    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $true
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Red"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "No VDA found."
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    elseif ($Total_Regist -eq 1)
						    {
							    $Regist_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $Regist_List_Datagrid.Add($Regist_list)
							    $SyncHash.datagrid_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_Maintenance_Registered.ItemsSource = $Regist_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.TextBox_Servers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Maintenance_Registered.text = $Simple_Regist_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Regist" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Grid_Simple_Maintenance_Registered.Visibility = "Visible"
									    $SyncHash.Refresh_Registration.Visibility = "Visible"
									    $SyncHash.Refresh_Registration_Simple.Visibility = "Visible"
									    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $true
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
						    else
						    {
							    $Regist_List_Datagrid = New-Object System.Collections.Generic.List[Object]
							    $Regist_List_Datagrid.AddRange($Regist_list)
							    $SyncHash.datagrid_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.datagrid_Maintenance_Registered.ItemsSource = $Regist_List_Datagrid }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.TextBox_Servers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_Servers_Maintenance_Registered.text = $Simple_Regist_List_String }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.TextBox_TotalServers_Maintenance_Registered.Dispatcher.Invoke([Action]{ $SyncHash.TextBox_TotalServers_Maintenance_Registered.text = "Total = $Total_Regist" }, [Windows.Threading.DispatcherPriority]::Normal)
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.Grid_Simple_Maintenance_Registered.Visibility = "Visible"
									    $SyncHash.Refresh_Registration.Visibility = "Visible"
									    $SyncHash.Refresh_Registration_Simple.Visibility = "Visible"
									    $SyncHash.Enable_Maintenance_MaintRegist.IsEnabled = $true
									    $SyncHash.MainLayer.IsEnabled = $true
								    }, "Normal")
						    }
					    }
					    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Registration_Simple_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
				    })
			    Worker
			    $Main_MB_Close.add_Click({
					    $Dialog_Main.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Registration_Simple " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Switch_MaintRegist_Full.add_Click({
		    $Grid_Simple_Maintenance_Registered.Visibility = "collapse"
		    $Grid_Detailled_Maintenance_Registered.Visibility = "visible"
	    })
    $Enable_Maintenance_MaintRegist.add_Click({
		    try
		    {
			    $VDAs = $datagrid_Maintenance_Registered.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to enable maintenance for the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "Enable_Maintenance_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_Maintenance_Registered.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Enable_Maintenance_MaintRegist " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Disble_Maintenance_MaintRegist.add_Click({
		    try
		    {
			    $VDAs = $datagrid_Maintenance_Registered.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $lastButtonClicked = $null
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to disable maintenance for the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "Disble_Maintenance_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_Maintenance_Registered.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Disble_Maintenance_MaintRegist " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOn_MaintRegist.add_Click({
		    try
		    {
			    $VDAs = $datagrid_Maintenance_Registered.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power on the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "PowerOn_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_Maintenance_Registered.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOn_MaintRegist " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $PowerOff_MaintRegist.add_Click({
		    try
		    {
			    $VDAs = $datagrid_Maintenance_Registered.SelectedItems
			    if ($VDAs.count -eq "0") { Show-Dialog_Main -Foreground "Red" -Text "Please select a VDA." }
			    Else
			    {
				    $MainLayer.IsEnabled = $false
				    $Main_TB_Confirm.Foreground = "Blue"
				    $Main_TB_Confirm.FontSize = "20"
				    $Main_TB_Confirm.text = "Are your sure you want to power off the selected VDAs ?"
				    $Dialog_Main_Confirm.IsOpen = $True
				    $global:Action = "PowerOff_AllVDAs"
				    $Main_MB_Confirm.add_Click({
						    $VDAs = $datagrid_Maintenance_Registered.SelectedItems
						    Main_MB_Confirm
					    })
				    $Main_MB_Cancel.add_Click({ Main_MB_Cancel })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_PowerOff_MaintRegist " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_Maintenance.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing VDAs in maintenance mode"
			    Refresh_Maintenance
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Maintenance " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Refresh_Registration.add_Click({
		    try
		    {
			    $Load_TB.Text = "Refreshing unregistered VDAs"
			    Refresh_Registration
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Refresh_Registration " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Export_MaintRegist.add_Click({
		    try
		    {
			    if ($datagrid_Maintenance_Registered.ItemsSource.count -eq 0) { Show-Dialog_Main -Foreground Red -Text "No VDAs to export." }
			    else
			    {
				    $FarmSelected = $S_Maintenance.selecteditem
				    $MainLayer.IsEnabled = $false
				    $SpinnerOverlayLayer_Main.Visibility = "Visible"
				    $Load_TB.Text = "Export in progress"
				    $Global:SyncHash_ExportVDAs = [hashtable]::Synchronized(@{
						    FarmSelected = $FarmSelected
						    ConfigPath   = $ConfigPath
					    })
				    $Runspace = [runspacefactory]::CreateRunspace()
				    $Runspace.ThreadOptions = "ReuseThread"
				    $Runspace.ApartmentState = "STA"
				    $Runspace.Open()
				    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
				    $Runspace.SessionStateProxy.SetVariable("SyncHash_ExportVDAs", $SyncHash_ExportVDAs)
				    $Worker = [PowerShell]::Create().AddScript({
						    try
						    {
							    $Farm = $SyncHash_ExportVDAs.FarmSelected
							    $date = get-date -Format MM_dd_yyyy
							    $Export_VDAs = $SyncHash_ExportVDAs.ConfigPath + "\Exports\Maintenance_$Farm-$date.xlsx"
							    $i = 0
							    while (Test-Path $Export_VDAs)
							    {
								    $i++
								    $Export_VDAs = $SyncHash_ExportVDAs.ConfigPath + "\Exports\Maintenance__$Farm-$date-$i.xlsx"
							    }
							    Import-Module ".\Configuration\PSExcel-master\PSExcel"
							    $SyncHash.datagrid_Maintenance_Registered.ItemsSource | Export-xlsx -Path $Export_VDAs -Table -Autofit
							    $SyncHash.Form.Dispatcher.Invoke([action]{
									    $SyncHash.SpinnerOverlayLayer_Main.Visibility = "Collapsed"
									    $SyncHash.MainLayer.IsEnabled = $false
									    $SyncHash.Main_MB.Foreground = "Blue"
									    $SyncHash.Main_MB.FontSize = "20"
									    $SyncHash.Main_MB.text = "You will find your extract here :`r`n$Export_VDAs"
									    $SyncHash.Dialog_Main.IsOpen = $True
								    }, "Normal")
						    }
						    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_MaintRegist_AddScript " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
					    })
				    Worker
				    $Main_MB_Close.add_Click({
						    $Dialog_Main.IsOpen = $False
						    $MainLayer.IsEnabled = $true
					    })
			    }
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Export_MaintRegist " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    ##############################
    # End_Maintenance_Registration
    ##############################
    ################
    # Start_Settings
    ################
    #### Theme
    $Toggle.add_Click({
		    try
		    {
			    $theme = [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::GetTheme($Form.Resources)
			    if ($Toggle.IsChecked -eq $true) { [MaterialDesignThemes.Wpf.ThemeExtensions]::SetBaseTheme($theme, [MaterialDesignThemes.Wpf.Theme]::Dark) }
			    if ($Toggle.IsChecked -eq $False) { [MaterialDesignThemes.Wpf.ThemeExtensions]::SetBaseTheme($theme, [MaterialDesignThemes.Wpf.Theme]::Light) }
			    [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::SetTheme($Form.Resources, $theme)
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Toggle " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Colors = @()
    $Colors = [System.Enum]::GetNames([MaterialDesignColors.PrimaryColor])
    foreach ($item in $Colors) { $S_Color.Items.Add($item) | Out-Null }
    $S_Color.Add_SelectionChanged({
		    try
		    {
			    $theme = [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::GetTheme($Form.Resources)
			    $Primary = [MaterialDesignColors.SwatchHelper]::Lookup[$S_Color.SelectedValue]
			    [MaterialDesignThemes.Wpf.ThemeExtensions]::SetPrimaryColor($theme, $Primary)
			    [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::SetTheme($Form.Resources, $theme)
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_S_Color " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    #### Configuration
    $Launch_configuration_Help.add_Click({ $Snackbar.MessageQueue.Enqueue("It will close XD Tool and launch the configuration window.") })
    $Launch_configuration.add_Click({
		    try
		    {
			    $MainLayer.IsEnabled = $false
			    $Main_TB_Confirm.Foreground = "Blue"
			    $Main_TB_Confirm.FontSize = "20"
			    $Main_TB_Confirm.text = "Are your sure you want to quit and launch the configuration window ?"
			    $Dialog_Main_Confirm.IsOpen = $True
			    $Main_MB_Confirm.add_Click({
					    $Dialog_Main_Confirm.IsOpen = $False
					    $MainLayer.IsEnabled = $true
					    $Form.Hide()
					    $Check_conf.content = "Configuration"
					    $Form_Check_conf.ShowDialog()
				    })
			    $Main_MB_Cancel.add_Click({
					    $Dialog_Main_Confirm.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Launch_configuration " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    #### Version
    $Changelog.add_Click({
		    try
		    {
			    $MainLayer.IsEnabled = $false
			    $Main_MB_Version.FontSize = "20"
			    $Main_MB_Version.text = "V2.0.6 : 09/27/2024"
			    $Main_MB_Version.Inlines.Add((New-Object System.Windows.Documents.Run -Property @{
						    Text	 = "`r`nLicenses part : Now all licenses types are displayed.
                                                                                                           `r`nV2.0.5 : 01/23/2024`r`nMinor bugs fixed.
                                                                                                           `r`nV2.0.4 : 12/14/2023`r`nUnderscore '_' and hyphen '-' are now accepted for DDC name in the configuration step.`r`nFixed CertificateVerificationFailed error when XD Tool retrieve licenses.
                                                                                                           `r`nV2.0.3 : 10/29/2023`r`nAdding farm with the DDC where XD Tool is running is now fixed.`r`nAdding farm by DDC with FQDN is now fixed.`r`nLoad config.xml already loaded is now blocked without crashing XD Tool.`r`nAdding new farm is now faster.
                                                                                                           `r`nV2.0.2 : 10/23/2023`r`n`Publications tab, part 'All publications', farms search fixed.
                                                                                                           `r`nV2.0.1 : 10/21/2023`r`n`Publications tab, part 'All publications', 'Access' column has been added.`r`nSome display issues and minor bugs fixed.
                                                                                                           `r`nV2.0.0 : 10/09/2023`r`nWindows Presentation Foundation has replaced Windows Forms.`r`nMaterial Design UI used.`r`nConfiguration of farms added."; FontSize = "14"
					    }))
			    $Dialog_Main_Version.IsOpen = $True
			    $Main_MB_Version_Close.add_Click({
					    $Dialog_Main_Version.IsOpen = $False
					    $MainLayer.IsEnabled = $true
				    })
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Changelog " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    #### Misc
    $Copy_contact.add_Click({
		    try
		    {
			    $Contact.content | Set-Clipboard
			    $Snackbar.MessageQueue.Enqueue("Contact copied to clipboard.")
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Copy_contact " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    $Copy_link.add_Click({
		    try
		    {
			    $Download_link.content | Set-Clipboard
			    $Snackbar.MessageQueue.Enqueue("Download link copied to clipboard.")
		    }
		    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Copy_link " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }
	    })
    ##############
    # End_Settings
    ##############
    try
    {
	
	    $Splash.Close()
	    if (-not (Test-Path $ConfigFile))
	    {
		    $Check_conf.content = "Configuration file has not been found."
		    $Form_Check_conf.ShowDialog()
	    }
	    Else
	    {
		    $datas = Import-Clixml -Path $ConfigFile
		    $Form.add_Loaded({
				    $Load_TB.Text = "Loading configuration file"
				    Process-FarmData -datas $datas
			    })
		    $Form.Add_Closing({
				    $Process = Get-Process XD_Tool -ErrorAction SilentlyContinue
				    if ($Process) { $Process | Stop-Process -Force }
			    })
		    $Form.ShowDialog() | Out-Null
	    }
    }
    catch { $_.Exception.Message + " // Line " + $_.InvocationInfo.ScriptLineNumber + " Part_Start " + " --> " + (get-date) | Out-File "$env:LOCALAPPDATA\XD_Tool\Log_Errors.txt" -Append }

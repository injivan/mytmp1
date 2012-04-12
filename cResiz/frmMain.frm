VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Auto Management -  [company name] - [app version]"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   1125
   ClientWidth     =   12060
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   12060
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Help?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   7560
      Width           =   735
   End
   Begin TabDlg.SSTab SSTabMain 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   706
      BackColor       =   0
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMain.frx":0576
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmMain.frx":0592
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameSearchResults"
      Tab(1).Control(1)=   "cmdAddtoOrder"
      Tab(1).Control(2)=   "cmdClearOrder(0)"
      Tab(1).Control(3)=   "cmdSaveOrder2"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "cmdManagePart"
      Tab(1).Control(6)=   "cmdSoldCash"
      Tab(1).Control(7)=   "cmdViewDetails"
      Tab(1).Control(8)=   "cmdScrap(1)"
      Tab(1).Control(9)=   "cmdMiscPart"
      Tab(1).Control(10)=   "cmdMiscRefund"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Transactions"
      TabPicture(2)   =   "frmMain.frx":05AE
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "SSTab_Invoicing"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Customer"
      TabPicture(3)   =   "frmMain.frx":05CA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab_Customer"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frmMain.frx":05E6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Timer AutoTimer 
         Interval        =   65000
         Left            =   -360
         Top             =   0
      End
      Begin VB.CommandButton cmdMiscRefund 
         Caption         =   "Misc Refund"
         Height          =   375
         Left            =   -71400
         TabIndex        =   42
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton cmdMiscPart 
         Caption         =   "Misc Part"
         Height          =   375
         Left            =   -70320
         TabIndex        =   43
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton cmdScrap 
         Caption         =   "Scrap Part"
         Height          =   375
         Index           =   1
         Left            =   -69120
         TabIndex        =   44
         Top             =   7080
         Width           =   1092
      End
      Begin VB.CommandButton cmdViewDetails 
         Caption         =   "Vehicle Details"
         Height          =   375
         Left            =   -73680
         TabIndex        =   41
         Top             =   7080
         Width           =   1332
      End
      Begin VB.CommandButton cmdSoldCash 
         Caption         =   "Sold Cash"
         Height          =   375
         Left            =   -64800
         TabIndex        =   49
         Top             =   7080
         Width           =   1092
      End
      Begin VB.CommandButton cmdManagePart 
         Caption         =   "Manage Part"
         Height          =   375
         Left            =   -74880
         TabIndex        =   40
         Top             =   7080
         Width           =   1212
      End
      Begin VB.Frame Frame3 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1572
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   11175
         Begin VB.TextBox txtSearchYear 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox txtCustomSearchString 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   31
            Top             =   1200
            Width           =   2172
         End
         Begin VB.ComboBox cboSiteList 
            Height          =   315
            ItemData        =   "frmMain.frx":32E8
            Left            =   5880
            List            =   "frmMain.frx":32EF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1200
            Width           =   1572
         End
         Begin VB.ComboBox cboSearchMode 
            Height          =   315
            ItemData        =   "frmMain.frx":32F8
            Left            =   4200
            List            =   "frmMain.frx":3308
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1200
            Width           =   1572
         End
         Begin VB.ComboBox cboCustomSearch 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1200
            Width           =   1692
         End
         Begin VB.CommandButton cmdResetSearch 
            Caption         =   "Reset"
            Height          =   375
            Left            =   7560
            TabIndex        =   34
            Top             =   1080
            Width           =   852
         End
         Begin VB.CommandButton cmdAutoRequest 
            Caption         =   "Auto Request"
            Height          =   375
            Left            =   8520
            TabIndex        =   35
            Top             =   1080
            Width           =   1212
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search Part(s)"
            Height          =   375
            Left            =   9720
            TabIndex        =   36
            Top             =   1080
            Width           =   1332
         End
         Begin VB.ComboBox cboPartsSearch 
            Height          =   315
            Left            =   8640
            Sorted          =   -1  'True
            TabIndex        =   29
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox cboPartTypeSearch 
            Height          =   315
            Left            =   6480
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cboMakeSearch 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox cboModelSearch 
            Height          =   315
            Left            =   4200
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Year"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   106
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Site Search"
            Height          =   252
            Index           =   31
            Left            =   5880
            TabIndex        =   63
            Top             =   960
            Width           =   1932
         End
         Begin VB.Label Label1 
            Caption         =   "Search Mode"
            Height          =   252
            Index           =   30
            Left            =   4200
            TabIndex        =   59
            Top             =   960
            Width           =   1932
         End
         Begin VB.Label Label1 
            Caption         =   "Search String"
            Height          =   252
            Index           =   23
            Left            =   1920
            TabIndex        =   57
            Top             =   960
            Width           =   3612
         End
         Begin VB.Label Label1 
            Caption         =   "Select or Enter Part"
            Height          =   255
            Index           =   16
            Left            =   8640
            TabIndex        =   56
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Part Type"
            Height          =   255
            Index           =   14
            Left            =   6480
            TabIndex        =   55
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Model"
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   54
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Custom Field "
            Height          =   252
            Index           =   43
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   1932
         End
         Begin VB.Label Label1 
            Caption         =   "Make"
            Height          =   255
            Index           =   24
            Left            =   1920
            TabIndex        =   4
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdSaveOrder2 
         Caption         =   "Save Order"
         Height          =   375
         Left            =   -66960
         TabIndex        =   46
         Top             =   7080
         Width           =   1092
      End
      Begin VB.CommandButton cmdClearOrder 
         Caption         =   "Clear Order"
         Height          =   375
         Index           =   0
         Left            =   -68040
         TabIndex        =   45
         Top             =   7080
         Width           =   1092
      End
      Begin VB.CommandButton cmdAddtoOrder 
         Caption         =   "Add to Order"
         Height          =   375
         Left            =   -65880
         TabIndex        =   48
         Top             =   7080
         Width           =   1092
      End
      Begin VB.Frame frameSearchResults 
         Caption         =   "Search Result(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4932
         Left            =   -74880
         TabIndex        =   5
         Top             =   2160
         Width           =   11175
         Begin MSFlexGridLib.MSFlexGrid grdSearchResults 
            Height          =   4215
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7435
            _Version        =   393216
            BackColor       =   8454143
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            SelectionMode   =   1
            AllowUserResizing=   3
         End
         Begin VB.TextBox txtCreateEbayHTMLTemplate 
            Height          =   975
            Left            =   7320
            MultiLine       =   -1  'True
            TabIndex        =   124
            Top             =   3120
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label lblAlertMessage 
            Alignment       =   2  'Center
            Caption         =   "alert message area"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   37
            Left            =   120
            TabIndex        =   104
            Top             =   4440
            Width           =   10815
         End
      End
      Begin TabDlg.SSTab SSTab_Invoicing 
         Height          =   6975
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12303
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Orders"
         TabPicture(0)   =   "frmMain.frx":3331
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame5"
         Tab(0).Control(1)=   "cmdClearOrder2"
         Tab(0).Control(2)=   "Frame11"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Transaction Reports"
         TabPicture(1)   =   "frmMain.frx":334D
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame12"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6492
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   11172
            Begin TabDlg.SSTab SSTab1 
               Height          =   5565
               Left            =   120
               TabIndex        =   121
               Top             =   840
               Width           =   10965
               _ExtentX        =   19341
               _ExtentY        =   9816
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   6
               TabHeight       =   520
               TabCaption(0)   =   "Detail"
               TabPicture(0)   =   "frmMain.frx":3369
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "gridInvoices"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Summary"
               TabPicture(1)   =   "frmMain.frx":3385
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "grdTranSummary"
               Tab(1).ControlCount=   1
               Begin MSFlexGridLib.MSFlexGrid gridInvoices 
                  Height          =   5175
                  Left            =   120
                  TabIndex        =   122
                  ToolTipText     =   "Double Click Invoice to View, Add a Payment or Print"
                  Top             =   360
                  Width           =   10815
                  _ExtentX        =   19076
                  _ExtentY        =   9128
                  _Version        =   393216
                  BackColor       =   8454143
                  BackColorBkg    =   14737632
                  SelectionMode   =   1
                  AllowUserResizing=   3
               End
               Begin MSFlexGridLib.MSFlexGrid grdTranSummary 
                  Height          =   5175
                  Left            =   -74880
                  TabIndex        =   123
                  ToolTipText     =   "Double Click Invoice to View, Add a Payment or Print"
                  Top             =   360
                  Width           =   10815
                  _ExtentX        =   19076
                  _ExtentY        =   9128
                  _Version        =   393216
                  BackColor       =   8454143
                  BackColorBkg    =   14737632
                  SelectionMode   =   1
                  AllowUserResizing=   3
               End
            End
            Begin VB.CommandButton cmdExportData 
               Caption         =   "Export"
               Height          =   375
               Left            =   9120
               TabIndex        =   118
               Top             =   360
               Width           =   735
            End
            Begin VB.ComboBox cboTransReport 
               Height          =   288
               ItemData        =   "frmMain.frx":33A1
               Left            =   120
               List            =   "frmMain.frx":33E1
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   480
               Width           =   4092
            End
            Begin VB.ComboBox cboCustomInvSearch 
               Height          =   315
               Left            =   7080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   480
               Width           =   1932
            End
            Begin VB.TextBox txtSearchInvoiceNo 
               Height          =   300
               Left            =   4560
               TabIndex        =   68
               ToolTipText     =   "Press Enter to Search"
               Top             =   480
               Width           =   2412
            End
            Begin VB.CommandButton cmdRefreshInv 
               Caption         =   "Refresh"
               Height          =   375
               Left            =   9960
               TabIndex        =   70
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   " By Field"
               Height          =   375
               Index           =   0
               Left            =   7080
               TabIndex        =   120
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Search       "
               Height          =   375
               Index           =   41
               Left            =   4560
               TabIndex        =   119
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label Label10 
               Caption         =   "Report"
               Height          =   252
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   2652
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Saved Order(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   -74880
            TabIndex        =   10
            Top             =   3720
            Width           =   10935
            Begin VB.CommandButton cmdSoldcashpendingorder 
               Caption         =   "Sold Cash"
               Height          =   375
               HelpContextID   =   202
               Left            =   3480
               TabIndex        =   117
               Top             =   2400
               Width           =   1215
            End
            Begin VB.CommandButton cmdSearchOrders 
               Caption         =   "Search"
               Height          =   375
               Left            =   9120
               TabIndex        =   115
               Top             =   2400
               Width           =   1575
            End
            Begin VB.TextBox txtOrderSearch 
               Height          =   300
               Left            =   6600
               TabIndex        =   114
               ToolTipText     =   "Press Enter to Search"
               Top             =   2400
               Width           =   2412
            End
            Begin VB.CommandButton cmdCancelOrder 
               Caption         =   "Cancel Order"
               Height          =   375
               Left            =   1800
               TabIndex        =   64
               Top             =   2400
               Width           =   1575
            End
            Begin VB.CommandButton cmdInvoiceOrder 
               Caption         =   "Invoice Order"
               Height          =   375
               Left            =   4800
               TabIndex        =   65
               Top             =   2400
               Width           =   1575
            End
            Begin VB.CommandButton cmdRefreshPendingOrders 
               Caption         =   "Refresh"
               Height          =   375
               Left            =   120
               TabIndex        =   62
               Top             =   2400
               Width           =   1575
            End
            Begin MSFlexGridLib.MSFlexGrid gridPendingOrders 
               Height          =   1935
               Left            =   120
               TabIndex        =   61
               ToolTipText     =   "Double Click Pending Order to View Order Contents"
               Top             =   360
               Width           =   10695
               _ExtentX        =   18865
               _ExtentY        =   3413
               _Version        =   393216
               BackColor       =   8454143
               BackColorBkg    =   14737632
               SelectionMode   =   1
            End
         End
         Begin VB.CommandButton cmdClearOrder2 
            Caption         =   "Clear Order"
            Height          =   375
            HelpContextID   =   200
            Left            =   -74760
            TabIndex        =   7
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Frame Frame5 
            Caption         =   "Order Contents"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   -74880
            TabIndex        =   6
            Top             =   360
            Width           =   10935
            Begin VB.CommandButton cmdSaveExistingOrder 
               Caption         =   "Save Existing Order"
               Height          =   375
               HelpContextID   =   202
               Left            =   6360
               TabIndex        =   116
               Top             =   2760
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.CommandButton cmdMiscPart2 
               Caption         =   "Misc Part"
               Height          =   375
               HelpContextID   =   203
               Left            =   5160
               TabIndex        =   60
               Top             =   2760
               Width           =   1092
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Sold Cash"
               Height          =   375
               HelpContextID   =   202
               Left            =   3480
               TabIndex        =   58
               Top             =   2760
               Width           =   1575
            End
            Begin VB.CommandButton cmdSaveOrder 
               Caption         =   "Save Order"
               Height          =   375
               HelpContextID   =   201
               Left            =   1800
               TabIndex        =   9
               Top             =   2760
               Width           =   1575
            End
            Begin MSFlexGridLib.MSFlexGrid gridOrderList 
               Height          =   2292
               Left            =   120
               TabIndex        =   66
               ToolTipText     =   "Double Click Part to Amend Nett Sale Price"
               Top             =   360
               Width           =   10692
               _ExtentX        =   18865
               _ExtentY        =   4048
               _Version        =   393216
               BackColor       =   8454143
               BackColorBkg    =   14737632
               SelectionMode   =   1
            End
         End
      End
      Begin TabDlg.SSTab SSTab_Customer 
         Height          =   6972
         Left            =   -75000
         TabIndex        =   2
         Top             =   480
         Width           =   11412
         _ExtentX        =   20135
         _ExtentY        =   12303
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Manage Customers"
         TabPicture(0)   =   "frmMain.frx":35DA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdRefreshCustLst"
         Tab(0).Control(1)=   "cmdDelCust"
         Tab(0).Control(2)=   "SearchString"
         Tab(0).Control(3)=   "Frame16"
         Tab(0).Control(4)=   "CustomerList"
         Tab(0).Control(5)=   "StatusBar2"
         Tab(0).Control(6)=   "Label1(35)"
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Customer Record"
         TabPicture(1)   =   "frmMain.frx":35F6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame15"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Enquires"
         TabPicture(2)   =   "frmMain.frx":3612
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "cmdRemoveAutoReq2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "cmdCheckArMatches"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "cmdRefreshCustAutoReqs"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "opCustAR"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "frameAROptions"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "cmdRefreshEnquiries"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "frameAR_ENQ"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "opCustEnquires"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "cmbEnqStatus"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).ControlCount=   9
         Begin VB.ComboBox cmbEnqStatus 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   480
            Width           =   3255
         End
         Begin VB.OptionButton opCustEnquires 
            Caption         =   "Customer Enquires"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   480
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.CommandButton cmdRefreshCustLst 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   -65280
            TabIndex        =   73
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelCust 
            Caption         =   "Delete"
            Height          =   375
            Left            =   -65280
            TabIndex        =   75
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox SearchString 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -74880
            TabIndex        =   71
            Top             =   840
            Width           =   6012
         End
         Begin VB.Frame Frame16 
            Caption         =   "Search Mode"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   -68640
            TabIndex        =   24
            Top             =   480
            Width           =   3132
            Begin VB.ComboBox cboCustSearchMode 
               Height          =   288
               Left            =   120
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   360
               Width           =   2892
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Customer Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6492
            Left            =   -74880
            TabIndex        =   12
            Top             =   360
            Width           =   11172
            Begin VB.ListBox lstPAF 
               Height          =   450
               Left            =   1680
               Sorted          =   -1  'True
               TabIndex        =   97
               Top             =   1680
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CommandButton cmdPAFLookup 
               Caption         =   "Find"
               Height          =   375
               Left            =   3720
               TabIndex        =   96
               Top             =   3120
               Width           =   975
            End
            Begin VB.TextBox txtCreditLimit 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   84
               Top             =   3480
               Width           =   2055
            End
            Begin VB.CommandButton cmdDelCust_Click2 
               Caption         =   "Delete"
               Height          =   375
               Left            =   6360
               TabIndex        =   92
               Top             =   5880
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtCustID 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   47
               Top             =   4080
               Width           =   2055
            End
            Begin VB.CommandButton cmdResetCustRec 
               Caption         =   "Reset"
               Height          =   375
               Left            =   7680
               TabIndex        =   91
               Top             =   5880
               Width           =   1455
            End
            Begin VB.TextBox NewCustomerRef 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   76
               Top             =   360
               Width           =   4095
            End
            Begin VB.CommandButton SaveCustomer 
               Caption         =   "Save"
               Height          =   375
               Left            =   9240
               TabIndex        =   90
               Top             =   5880
               Width           =   1692
            End
            Begin VB.TextBox NewCustomerTel 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7440
               MaxLength       =   50
               TabIndex        =   85
               Top             =   360
               Width           =   3255
            End
            Begin VB.TextBox NewCustomerPostcode 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   83
               Top             =   3120
               Width           =   2055
            End
            Begin VB.TextBox NewCustomerTown 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   81
               Top             =   2400
               Width           =   4095
            End
            Begin VB.TextBox NewCustomerAdd3 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   82
               Top             =   2760
               Width           =   4095
            End
            Begin VB.TextBox NewCustomerAdd2 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   80
               Top             =   2040
               Width           =   4095
            End
            Begin VB.TextBox NewCustomerAdd1 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   79
               Top             =   1680
               Width           =   4095
            End
            Begin VB.TextBox NewCustomerName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   78
               Top             =   1200
               Width           =   4095
            End
            Begin VB.TextBox NewCustomerEmail 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7440
               MaxLength       =   50
               TabIndex        =   88
               Top             =   1440
               Width           =   3255
            End
            Begin VB.TextBox NewCustomerComments 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1410
               Left            =   7440
               MaxLength       =   50
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   89
               Top             =   1920
               Width           =   3255
            End
            Begin VB.TextBox NewCustomerBusName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   77
               Top             =   840
               Width           =   4095
            End
            Begin VB.TextBox NewCustomerFax 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7440
               MaxLength       =   50
               TabIndex        =   86
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox NewCustomerMob 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7440
               MaxLength       =   50
               TabIndex        =   87
               Top             =   1080
               Width           =   3255
            End
            Begin VB.Label Label5 
               Caption         =   "Credit Limit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   120
               TabIndex        =   53
               Top             =   3480
               Width           =   1452
            End
            Begin VB.Label Label34 
               Caption         =   "Customer ID"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   120
               TabIndex        =   50
               Top             =   4080
               Width           =   1452
            End
            Begin VB.Label Label33 
               Caption         =   "Customer Ref"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label Label32 
               Caption         =   "Telephone No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6000
               TabIndex        =   22
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label31 
               Caption         =   "Postcode"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   3120
               Width           =   1455
            End
            Begin VB.Label Label30 
               Caption         =   "Customer Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   1200
               Width           =   2175
            End
            Begin VB.Label Label29 
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label28 
               Caption         =   "Town"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Email Address"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   34
               Left            =   6000
               TabIndex        =   17
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Label Label27 
               Caption         =   "Comments / Details"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6000
               TabIndex        =   16
               Top             =   1920
               Width           =   1455
            End
            Begin VB.Label Label26 
               Caption         =   "Business Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   840
               Width           =   2175
            End
            Begin VB.Label Label25 
               Caption         =   "Fax No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6000
               TabIndex        =   14
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label24 
               Caption         =   "Mobile No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6000
               TabIndex        =   13
               Top             =   1080
               Width           =   1575
            End
         End
         Begin VB.Frame frameAR_ENQ 
            Caption         =   "Customer Enquiries"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6015
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   11055
            Begin MSFlexGridLib.MSFlexGrid grdCustAutoReqs 
               Height          =   5655
               Left            =   120
               TabIndex        =   93
               Top             =   240
               Width           =   10815
               _ExtentX        =   19076
               _ExtentY        =   9975
               _Version        =   393216
               BackColor       =   8454143
               BackColorBkg    =   14737632
               SelectionMode   =   1
            End
         End
         Begin MSFlexGridLib.MSFlexGrid CustomerList 
            Height          =   5052
            Left            =   -74880
            TabIndex        =   74
            ToolTipText     =   "Double Click Customer Record to View or Amend"
            Top             =   1440
            Width           =   11052
            _ExtentX        =   19500
            _ExtentY        =   8916
            _Version        =   393216
            Rows            =   1
            Cols            =   9
            BackColor       =   8454143
            BackColorBkg    =   16777215
            AllowUserResizing=   3
         End
         Begin MSComctlLib.StatusBar StatusBar2 
            Height          =   375
            Left            =   -75000
            TabIndex        =   37
            Top             =   8070
            Width           =   10770
            _ExtentX        =   18997
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdRefreshEnquiries 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   9840
            TabIndex        =   125
            Top             =   360
            Width           =   1212
         End
         Begin VB.Frame frameAROptions 
            Caption         =   "Options"
            Height          =   855
            Left            =   1560
            TabIndex        =   107
            Top             =   1920
            Visible         =   0   'False
            Width           =   8655
            Begin VB.ComboBox Combo2 
               Height          =   315
               ItemData        =   "frmMain.frx":362E
               Left            =   3840
               List            =   "frmMain.frx":3638
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   110
               Top             =   360
               Width           =   1815
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "frmMain.frx":3653
               Left            =   840
               List            =   "frmMain.frx":3660
               Style           =   2  'Dropdown List
               TabIndex        =   109
               Top             =   360
               Width           =   1815
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Refresh"
               Height          =   375
               Left            =   6720
               TabIndex        =   108
               Top             =   360
               Width           =   1212
            End
            Begin VB.Label Label3 
               Caption         =   "Order "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3120
               TabIndex        =   112
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Sort By:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   39
               Left            =   120
               TabIndex        =   111
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.OptionButton opCustAR 
            Caption         =   "Customer Auto Requests"
            Height          =   255
            Left            =   2520
            TabIndex        =   98
            Top             =   2640
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton cmdRefreshCustAutoReqs 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   7800
            TabIndex        =   95
            Top             =   4800
            Width           =   1095
         End
         Begin VB.CommandButton cmdCheckArMatches 
            Caption         =   "Check Auto Requests"
            Height          =   375
            Left            =   5880
            TabIndex        =   113
            Top             =   4920
            Width           =   2175
         End
         Begin VB.CommandButton cmdRemoveAutoReq2 
            Caption         =   "Remove  Auto Request"
            Height          =   375
            Left            =   5160
            TabIndex        =   94
            Top             =   4920
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Enter Search:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   -74880
            TabIndex        =   38
            Top             =   600
            Width           =   7215
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   101
      Top             =   7560
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2011-08-31 ã."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "18:05"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   103
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   5160
      TabIndex        =   102
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu Tools 
      Caption         =   "Tools"
      Begin VB.Menu mm_1stchoice_retrievequote 
         Caption         =   "Retrieve Quote"
         Visible         =   0   'False
      End
      Begin VB.Menu mm_firstchoice_blankline 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu reglookup 
         Caption         =   "Vehicle Registration Lookup"
      End
      Begin VB.Menu SendEmail_tool 
         Caption         =   "Send Email"
      End
      Begin VB.Menu sendsms 
         Caption         =   "Send SMS"
      End
      Begin VB.Menu ChangeUserPassword 
         Caption         =   "Change User Password"
      End
      Begin VB.Menu WebUpload 
         Caption         =   "Web Upload"
      End
      Begin VB.Menu menu_import 
         Caption         =   "Import"
         Begin VB.Menu importSales 
            Caption         =   "Sales"
         End
         Begin VB.Menu mm_downloadprocess_paypalsales 
            Caption         =   "Download / Process Paypal Sales"
         End
      End
      Begin VB.Menu export 
         Caption         =   "Export"
         Begin VB.Menu NotofDestruction 
            Caption         =   "Notification of Destruction"
         End
         Begin VB.Menu groupdataexport 
            Caption         =   "Group Data Export"
         End
         Begin VB.Menu Export_ebay 
            Caption         =   "Ebay"
         End
         Begin VB.Menu ebay_batch_validation 
            Caption         =   "Ebay Batch Validation"
         End
         Begin VB.Menu ebay_batchupload 
            Caption         =   "Ebay Batch Upload"
         End
      End
      Begin VB.Menu NewStatement 
         Caption         =   "New Statement"
      End
      Begin VB.Menu DatabaseQuery 
         Caption         =   "Database"
         Begin VB.Menu DBQuery 
            Caption         =   "Query"
         End
         Begin VB.Menu DBManagement 
            Caption         =   "Management (Lyons Systems Only)"
         End
         Begin VB.Menu db_exportSQL_scripts 
            Caption         =   "Export Access to SQL"
         End
      End
      Begin VB.Menu settings_old 
         Caption         =   "Settings"
         Begin VB.Menu mm_allsettings 
            Caption         =   "All Settings"
         End
         Begin VB.Menu mm_settings_vendor 
            Caption         =   "Vendors"
         End
         Begin VB.Menu imageconverter 
            Caption         =   "Image Converter"
         End
      End
      Begin VB.Menu UpdateApp 
         Caption         =   "Update Application"
      End
      Begin VB.Menu MENBackup 
         Caption         =   "Backup"
         Begin VB.Menu RunBackup 
            Caption         =   "Run Backup"
         End
         Begin VB.Menu RemoveableBackup 
            Caption         =   "Removeable Media Backup"
         End
      End
      Begin VB.Menu reports 
         Caption         =   "Reports"
      End
      Begin VB.Menu greenreport 
         Caption         =   "Green Parts Report"
      End
      Begin VB.Menu workoffline 
         Caption         =   "Work Offline"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mm_inventory 
      Caption         =   "Inventory"
   End
   Begin VB.Menu blank1 
      Caption         =   " "
   End
   Begin VB.Menu blank2 
      Caption         =   ""
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu HelpGuide 
         Caption         =   "Help Guide"
      End
      Begin VB.Menu helpfor 
         Caption         =   "Help with..."
         Index           =   0
      End
      Begin VB.Menu Support 
         Caption         =   "Support"
      End
   End
   Begin VB.Menu menuPopUpMenus_search 
      Caption         =   "PopUpMenus-SEARCH"
      Visible         =   0   'False
      Begin VB.Menu search_menu_ar_QuoteReply 
         Caption         =   "Quote Reply"
         Visible         =   0   'False
      End
      Begin VB.Menu search_menu_ar_CancelRequest 
         Caption         =   "Cancel Request"
         Visible         =   0   'False
      End
      Begin VB.Menu search_menu_ar_separator 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu export_createbatch 
         Caption         =   "Create Export Batch"
         Visible         =   0   'False
      End
      Begin VB.Menu search_menu_ar_separator2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuAddOrder 
         Caption         =   "Add to Order"
      End
      Begin VB.Menu menuscrap 
         Caption         =   "Scrap"
      End
      Begin VB.Menu menuManage 
         Caption         =   "Manage"
      End
      Begin VB.Menu menuVehdetails 
         Caption         =   "Vehicle Details"
      End
      Begin VB.Menu Stockdetails 
         Caption         =   "Stock Details"
      End
      Begin VB.Menu menuSaveOrder 
         Caption         =   "Save Order and Invoice"
      End
      Begin VB.Menu menuCashsale 
         Caption         =   "Cash Sale"
      End
      Begin VB.Menu menuMiscPart 
         Caption         =   "Add Misc Part"
      End
      Begin VB.Menu menuMiscRefund 
         Caption         =   "Misc Refund"
      End
      Begin VB.Menu menuClearOrder 
         Caption         =   "Clear Order"
      End
   End
   Begin VB.Menu menuPopUpMenus_nettsale 
      Caption         =   "menuPopUpMenus_nettsale"
      Visible         =   0   'False
      Begin VB.Menu menuPriceSchedule 
         Caption         =   "Pricing Schedule"
      End
      Begin VB.Menu suggestprice 
         Caption         =   "Suggest Price"
      End
   End
   Begin VB.Menu menuPopUpMenus_orders 
      Caption         =   "menuPopUpMenus_orders"
      Visible         =   0   'False
      Begin VB.Menu menuViewOrder 
         Caption         =   "View Order"
      End
      Begin VB.Menu menuInvoiceOrder 
         Caption         =   "Invoice Order"
      End
      Begin VB.Menu menusoldcash 
         Caption         =   "Sold Cash"
      End
      Begin VB.Menu menuCancelOrder 
         Caption         =   "Cancel Order"
      End
      Begin VB.Menu contactcustemail_order 
         Caption         =   "Contact Customer by Email"
      End
      Begin VB.Menu contactcustsms_order 
         Caption         =   "Contact Customer by SMS"
      End
      Begin VB.Menu order_paypalmoneyrequest 
         Caption         =   "Paypal Money Request"
      End
   End
   Begin VB.Menu menuInvoicesDelivery 
      Caption         =   "menuInvoicesDelivery"
      Visible         =   0   'False
      Begin VB.Menu createdelivery 
         Caption         =   "Create Delivery"
      End
      Begin VB.Menu inv_markdispatched 
         Caption         =   "Mark as Dispatched"
      End
      Begin VB.Menu inv_printselection 
         Caption         =   "Print (Selection)"
      End
      Begin VB.Menu contactcustemail_invoice 
         Caption         =   "Contact Customer by Email"
      End
      Begin VB.Menu contactcustsms_invoice 
         Caption         =   "Contact Customer by SMS"
      End
      Begin VB.Menu chaseos_email 
         Caption         =   "Chase Outstanding Balance by Email"
      End
      Begin VB.Menu chaseos_sms 
         Caption         =   "Chase Outstanding Balance by SMS"
      End
   End
   Begin VB.Menu menuPopupAutoRequests 
      Caption         =   "menuPopupAutoRequests"
      Visible         =   0   'False
      Begin VB.Menu email_matchedAR 
         Caption         =   "Notify Customer (by Email)"
      End
      Begin VB.Menu email_notifyall 
         Caption         =   "Notify All Customers (by Email)"
      End
      Begin VB.Menu sms_matchedAR 
         Caption         =   "Notify Customer (by SMS)"
      End
      Begin VB.Menu sms_notifyall 
         Caption         =   "Notify All Customers (by SMS)"
      End
   End
   Begin VB.Menu Customer_ARs 
      Caption         =   "Customer_ARs"
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      Begin VB.Menu CustAR_Email 
         Caption         =   "Contact Customer by Email"
      End
      Begin VB.Menu CustAR_SMS 
         Caption         =   "Contact Customer by SMS"
      End
   End
   Begin VB.Menu workingoffline_indicator 
      Caption         =   "         **OFFLINE MODE**"
      Visible         =   0   'False
   End
   Begin VB.Menu CheckAROptionsMenu 
      Caption         =   "CheckARVendorOption"
      Visible         =   0   'False
      Begin VB.Menu checkautorequests 
         Caption         =   "Check Auto Requests"
      End
      Begin VB.Menu checkvendorrequests 
         Caption         =   "Check 1st Choice Requests"
      End
   End
   Begin VB.Menu menuAwaitingQuote 
      Caption         =   "1stChoiceAwaitingQuote"
      Visible         =   0   'False
      Begin VB.Menu firstchoice_QuoteReply 
         Caption         =   "Quote Reply"
      End
      Begin VB.Menu firstchoice_viewpartdetails 
         Caption         =   "View Part Details"
      End
      Begin VB.Menu firstchoice_cancelmatch 
         Caption         =   "Cancel Request"
      End
   End
   Begin VB.Menu menuAllUnmatchedQuote 
      Caption         =   "1stChoice-UnmatchedQuote"
      Visible         =   0   'False
      Begin VB.Menu firstchoice_QuoteReplyUnmatched 
         Caption         =   "Quote Reply"
      End
      Begin VB.Menu firstchoice_cancelmatch_otherstatus 
         Caption         =   "Cancel Request"
      End
   End
   Begin VB.Menu SearchOptions 
      Caption         =   "SearchOptions"
      Visible         =   0   'False
      Begin VB.Menu search_sortresults 
         Caption         =   "Sort By..."
      End
      Begin VB.Menu search_groupby_parttype 
         Caption         =   "Group By:  Part Type"
      End
      Begin VB.Menu search_partlinks 
         Caption         =   "Part Links"
      End
      Begin VB.Menu groupbyvehicles 
         Caption         =   "Vehicles"
      End
      Begin VB.Menu ebay_batch_edit 
         Caption         =   "Ebay Batch Edit"
      End
   End
   Begin VB.Menu InvoicePrintOptions 
      Caption         =   "InvoicePrint"
      Visible         =   0   'False
      Begin VB.Menu printworkorder 
         Caption         =   "Print Work Order"
      End
   End
   Begin VB.Menu createdelmenu 
      Caption         =   "CreateDel"
      Visible         =   0   'False
      Begin VB.Menu del_amendcourier 
         Caption         =   "Amend Courier"
      End
   End
   Begin VB.Menu ImageAttachmm 
      Caption         =   "ImageAttacher"
      Visible         =   0   'False
      Begin VB.Menu imageatt_allocateimage 
         Caption         =   "Allocate Image to Part(s)"
      End
   End
   Begin VB.Menu InvoiceNettGrosspop 
      Caption         =   "InvoiceNettGross"
      Visible         =   0   'False
      Begin VB.Menu changenettgrossmp 
         Caption         =   "changenettgross"
      End
   End
   Begin VB.Menu picklistmenu 
      Caption         =   "PickList"
      Visible         =   0   'False
      Begin VB.Menu pick_list 
         Caption         =   "Print Pick List"
      End
   End
   Begin VB.Menu editebayoptions 
      Caption         =   "Edit Ebay"
      Visible         =   0   'False
      Begin VB.Menu editebay_field 
         Caption         =   "Edit Ebay Field(s)"
      End
      Begin VB.Menu ebay_managepart 
         Caption         =   "Manage"
      End
      Begin VB.Menu ebay_batchuploadchanges 
         Caption         =   "Upload Changes to Ebay"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim sR As cResaze

Private Sub Form_Initialize()
    Set sR = New cResaze
    sR.GetFrm Me, 7935, 12180
    
    '===========================================
    '  7935 = Me.ScaleHeight
    ' 12180 = Me.ScaleWidth
    '
    'DO NOT CALL THE FUNCTION like that
    ' sR.GetFrm Me, Me.ScaleHeight, Me.ScaleWidth
    '
    'BECAUSE It will DON'T Work properly
    '
    '======================================
End Sub


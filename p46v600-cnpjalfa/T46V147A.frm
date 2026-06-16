VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V147A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proposta Suspensa - Problemas na digitação"
   ClientHeight    =   8100
   ClientLeft      =   570
   ClientTop       =   1650
   ClientWidth     =   10740
   Icon            =   "T46V147A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraDocto 
      BackColor       =   &H80000018&
      Caption         =   "Documentos Faltantes"
      ClipControls    =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   3120
      Left            =   2760
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   5685
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Telefone"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   16
         Left            =   2745
         TabIndex        =   18
         Top             =   780
         Width           =   2130
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "DDD"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   15
         Left            =   270
         TabIndex        =   17
         Top             =   780
         Width           =   2130
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "RG (Numero, Data e/ou Orgão de Expedição)"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   16
         Top             =   240
         Width           =   4200
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Ramo de Atividade"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   2745
         TabIndex        =   15
         Top             =   510
         Width           =   2310
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Razão Social"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   14
         Top             =   510
         Width           =   1455
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Complemento do Segurado"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   13
         Top             =   1350
         Width           =   2310
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Cidade do Segurado"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   270
         TabIndex        =   12
         Top             =   1905
         Width           =   1815
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Bairro do Segurado"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   11
         Top             =   1635
         Width           =   1860
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Endereço do Segurado"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   10
         Top             =   1065
         Width           =   2040
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Cep do Segurado"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   13
         Left            =   270
         TabIndex        =   9
         Top             =   2475
         Width           =   1815
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Bairro de Correspondência"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   2745
         TabIndex        =   8
         Top             =   1635
         Width           =   2265
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Cidade de Correspondência"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   2745
         TabIndex        =   7
         Top             =   1905
         Width           =   2400
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Complemento de Correspondência"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   2745
         TabIndex        =   6
         Top             =   1350
         Width           =   2805
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Endereço de Correspondência"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   2745
         TabIndex        =   5
         Top             =   1065
         Width           =   2535
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Cep de Correspondência"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   14
         Left            =   2745
         TabIndex        =   4
         Top             =   2475
         Width           =   2130
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Estado de Correspondência"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   12
         Left            =   2745
         TabIndex        =   3
         Top             =   2190
         Width           =   2355
      End
      Begin VB.CheckBox ChkDocto 
         BackColor       =   &H80000018&
         Caption         =   "Estado do Segurado"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   11
         Left            =   270
         TabIndex        =   2
         Top             =   2190
         Width           =   2310
      End
      Begin VB.CommandButton CmdDoctoFec 
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4635
         TabIndex        =   1
         Top             =   2745
         Width           =   960
      End
   End
   Begin VB.Frame FraItens 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4035
      Left            =   0
      TabIndex        =   64
      Top             =   4080
      Width           =   10695
      Begin VB.CheckBox ChkItem 
         Caption         =   "Perfil do Segurado alterado, enviar novo perfil"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   211
         Tag             =   "9025"
         Top             =   1605
         Width           =   4335
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Nota Fiscal Ilegível"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   210
         Tag             =   "9003"
         Top             =   120
         Width           =   4335
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Nota Fiscal sem data de saída do veículo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   209
         Tag             =   "9004"
         Top             =   705
         Width           =   4335
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Nota Fiscal para veículo 0KM - Está faltando"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   208
         Tag             =   "9005"
         Top             =   405
         Width           =   4335
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Perfil do Segurado - Dados divegentes"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   207
         Tag             =   "9010"
         Top             =   1005
         Width           =   3030
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Perfil do Segurado - Assinaturas Faltantes"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   206
         Tag             =   "9011"
         Top             =   1305
         Width           =   3240
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Veículos Importados (DI ou DUT) - enviar cópia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   205
         Tag             =   "9015"
         Top             =   2205
         Width           =   4335
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Placa e/ou Chassi - duplicidade - Apólice:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   204
         Tag             =   "9017"
         Top             =   2505
         Width           =   3285
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Placa e/ou Chassi - Informar"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   203
         Tag             =   "9049"
         Top             =   2805
         Width           =   2715
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Informar modelo completo do veículo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   135
         TabIndex        =   202
         Tag             =   "9019"
         Top             =   3105
         Width           =   4335
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Informar modelo completo do acessório"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   135
         TabIndex        =   201
         Tag             =   "9020"
         Top             =   3405
         Width           =   4335
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   7
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   200
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   7
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   199
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   7
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   198
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   7
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   197
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   7
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   196
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   7
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   195
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   7
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   194
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   7
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   193
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   7
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   192
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   2
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   191
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   2
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   190
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   2
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   189
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   2
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   188
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   2
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   187
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   2
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   186
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   2
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   185
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   2
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   184
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   2
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   183
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   1
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   182
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   1
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   181
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   1
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   180
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   1
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   179
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   1
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   178
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   1
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   177
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   1
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   176
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   1
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   175
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   1
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   174
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   0
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   173
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   0
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   172
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   0
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   171
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   0
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   170
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   0
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   169
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   0
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   168
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   0
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   167
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   0
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   166
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   0
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   165
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   11
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   164
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   11
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   163
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   11
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   162
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   11
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   161
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   11
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   160
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   11
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   159
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   11
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   158
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   11
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   157
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   10
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   156
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   10
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   155
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   10
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   154
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   10
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   153
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   10
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   152
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   10
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   151
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   10
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   150
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   10
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   149
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   10
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   148
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   9
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   147
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   9
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   146
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   9
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   145
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   9
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   144
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   9
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   143
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   9
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   142
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   9
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   141
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   9
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   140
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   9
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   139
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   8
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   138
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   8
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   137
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   8
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   136
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   8
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   135
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   8
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   134
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   8
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   133
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   8
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   132
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   8
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   131
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   8
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   130
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   6
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   129
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   6
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   128
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   6
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   127
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   6
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   126
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   6
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   125
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   6
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   124
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   6
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   123
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   6
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   122
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   6
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   121
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   5
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   120
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   5
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   119
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   5
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   118
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   5
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   117
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   5
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   116
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   5
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   115
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   5
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   114
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   5
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   113
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   5
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   112
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   4
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   111
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   4
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   110
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   4
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   109
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   4
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   108
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   4
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   107
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   4
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   106
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   4
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   105
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   4
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   104
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   4
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   103
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   3
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   102
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   3
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   101
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   3
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   100
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   3
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   99
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   3
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   98
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   3
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   97
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   3
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   96
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   3
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   95
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   3
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   94
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   12
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   93
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item3 
         Height          =   285
         Index           =   12
         Left            =   6492
         MaxLength       =   4
         TabIndex        =   92
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item4 
         Height          =   285
         Index           =   12
         Left            =   6993
         MaxLength       =   4
         TabIndex        =   91
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item5 
         Height          =   285
         Index           =   12
         Left            =   7494
         MaxLength       =   4
         TabIndex        =   90
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item6 
         Height          =   285
         Index           =   12
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   89
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item7 
         Height          =   285
         Index           =   12
         Left            =   8496
         MaxLength       =   4
         TabIndex        =   88
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item8 
         Height          =   285
         Index           =   12
         Left            =   8997
         MaxLength       =   4
         TabIndex        =   87
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item9 
         Height          =   285
         Index           =   12
         Left            =   9498
         MaxLength       =   4
         TabIndex        =   86
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item10 
         Height          =   285
         Index           =   12
         Left            =   10005
         MaxLength       =   4
         TabIndex        =   85
         Top             =   1890
         Width           =   495
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Perfil Faltante"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   135
         TabIndex        =   84
         Tag             =   "9038"
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtDados3 
         Height          =   285
         Left            =   3300
         MaxLength       =   240
         TabIndex        =   83
         Top             =   980
         Width           =   2070
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   12
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   82
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   3
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   81
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   4
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   80
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   5
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   79
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   6
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   78
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   8
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   77
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   9
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   76
         Top             =   3090
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   10
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   75
         Top             =   3390
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   0
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   74
         Top             =   105
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   1
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   73
         Top             =   405
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   2
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   72
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   7
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   71
         Top             =   2490
         Width           =   495
      End
      Begin VB.TextBox txtApolice7 
         Height          =   285
         Left            =   3495
         MaxLength       =   240
         TabIndex        =   69
         Top             =   2490
         Width           =   1860
      End
      Begin VB.CheckBox ChkItem 
         Caption         =   "Veículo Valor Determinado"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   11
         Left            =   135
         TabIndex        =   68
         Tag             =   "9044"
         Top             =   3735
         Width           =   4335
      End
      Begin VB.TextBox Item1 
         Height          =   285
         Index           =   11
         Left            =   5490
         MaxLength       =   4
         TabIndex        =   67
         Top             =   3690
         Width           =   495
      End
      Begin VB.TextBox Item2 
         Height          =   285
         Index           =   11
         Left            =   5991
         MaxLength       =   4
         TabIndex        =   66
         Top             =   3675
         Width           =   495
      End
      Begin VB.ComboBox CmbTipInformar8 
         Height          =   315
         ItemData        =   "T46V147A.frx":000C
         Left            =   3495
         List            =   "T46V147A.frx":001C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2790
         Width           =   1845
      End
      Begin VB.ComboBox CmbAssPerfil4 
         Height          =   315
         ItemData        =   "T46V147A.frx":0041
         Left            =   3420
         List            =   "T46V147A.frx":0051
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1305
         Width           =   1950
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   0
         X2              =   5400
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line1 
         X1              =   5415
         X2              =   5415
         Y1              =   120
         Y2              =   3990
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   0
         X2              =   5415
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   0
         X2              =   5415
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   0
         X2              =   5370
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   0
         X2              =   5415
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Line Line2 
         Index           =   5
         X1              =   0
         X2              =   5415
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line2 
         Index           =   9
         X1              =   0
         X2              =   5415
         Y1              =   3075
         Y2              =   3075
      End
      Begin VB.Line Line2 
         Index           =   10
         X1              =   0
         X2              =   5415
         Y1              =   3375
         Y2              =   3375
      End
      Begin VB.Line Line2 
         Index           =   12
         X1              =   15
         X2              =   5415
         Y1              =   3975
         Y2              =   3975
      End
      Begin VB.Line Line2 
         Index           =   6
         X1              =   0
         X2              =   5415
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Line Line2 
         Index           =   7
         X1              =   0
         X2              =   5430
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line2 
         Index           =   8
         X1              =   0
         X2              =   5415
         Y1              =   2775
         Y2              =   2775
      End
      Begin VB.Line Line2 
         Index           =   11
         X1              =   0
         X2              =   5415
         Y1              =   3675
         Y2              =   3675
      End
      Begin VB.Line Line2 
         Index           =   14
         X1              =   0
         X2              =   5415
         Y1              =   85
         Y2              =   85
      End
   End
   Begin VB.TextBox TxtNumPI 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   705
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   105
      Width           =   1440
   End
   Begin VB.TextBox txtObservacao 
      Height          =   510
      Left            =   3285
      MaxLength       =   240
      MultiLine       =   -1  'True
      TabIndex        =   62
      Top             =   0
      Width           =   6150
   End
   Begin VB.Frame FraOutrosItens 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   945
      Left            =   0
      TabIndex        =   45
      Top             =   2925
      Width           =   10695
      Begin VB.CheckBox ChkOutrosItems 
         Caption         =   "Devido a vigência parcelamento em         vezes - Valor:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   57
         Tag             =   "9039"
         Top             =   390
         Width           =   4245
      End
      Begin VB.CheckBox ChkOutrosItems 
         Caption         =   "Cálculo não confere"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Tag             =   "9023"
         Top             =   120
         Width           =   1785
      End
      Begin VB.CheckBox ChkOutrosItems 
         Caption         =   "Apólice anterior da congênere"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   5490
         TabIndex        =   55
         Tag             =   "9013"
         Top             =   90
         Width           =   2475
      End
      Begin VB.TextBox txtValor1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4470
         TabIndex        =   54
         Top             =   390
         Width           =   840
      End
      Begin VB.CheckBox ChkOutrosItems 
         Caption         =   "Falta Cálculo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   5490
         TabIndex        =   52
         Tag             =   "9027"
         Top             =   360
         Width           =   1380
      End
      Begin VB.CheckBox ChkOutrosItems 
         Caption         =   "Pagto antecipado divergente - Parcela Paga R$"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   51
         Tag             =   "9033"
         Top             =   690
         Width           =   3750
      End
      Begin VB.TextBox txtPremio0 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2595
         TabIndex        =   50
         Top             =   90
         Width           =   1005
      End
      Begin VB.TextBox txtParc0 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4470
         TabIndex        =   49
         Top             =   90
         Width           =   840
      End
      Begin VB.TextBox txtPgAntItem2 
         Height          =   285
         Left            =   3870
         MaxLength       =   240
         TabIndex        =   48
         Top             =   675
         Width           =   1440
      End
      Begin VB.ComboBox CmbApolice3 
         Height          =   315
         ItemData        =   "T46V147A.frx":0072
         Left            =   8055
         List            =   "T46V147A.frx":007F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   45
         Width           =   2550
      End
      Begin VB.CheckBox ChkOutrosItems 
         Caption         =   "Enviar ao SAC"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   5
         Left            =   5490
         TabIndex        =   46
         Tag             =   "9041"
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox txtParc1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3075
         TabIndex        =   53
         Top             =   390
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "Prêmio"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   2010
         TabIndex        =   59
         Top             =   90
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Parcela:"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   3780
         TabIndex        =   58
         Top             =   90
         Width           =   600
      End
   End
   Begin VB.Frame FraComum 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2535
      Left            =   0
      TabIndex        =   19
      Top             =   450
      Width           =   10695
      Begin VB.CheckBox ChkComum 
         Caption         =   "Rasuras na Proposta"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   20
         Left            =   3480
         TabIndex        =   215
         Tag             =   "9060"
         Top             =   352
         Width           =   1890
      End
      Begin VB.TextBox txtSituac11 
         Height          =   285
         Left            =   8385
         MaxLength       =   240
         TabIndex        =   31
         Top             =   870
         Width           =   2220
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Falta parcelamento"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   44
         Tag             =   "9002"
         Top             =   75
         Width           =   1725
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Informar corretor e código"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   43
         Tag             =   "9007"
         Top             =   352
         Width           =   2745
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Débito em c/c sem declaração assinada pelo segurado"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   42
         Tag             =   "9009"
         Top             =   1460
         Width           =   4500
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Corretor desativado"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   41
         Tag             =   "9008"
         Top             =   629
         Width           =   1740
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Conta Corrente inválida"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   40
         Tag             =   "9030"
         Top             =   1183
         Width           =   2115
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "CPF/CNPJ incorreto ou faltante"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   39
         Tag             =   "9016"
         Top             =   906
         Width           =   2595
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Falta Comissão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   38
         Tag             =   "9031"
         Top             =   1737
         Width           =   1500
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Pagto antecipado divergente - Parcela Paga R$"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   5490
         TabIndex        =   37
         Tag             =   "9033"
         Top             =   365
         Width           =   3735
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Apólice pendente de pagamento (superior a 45 dias)"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   5490
         TabIndex        =   36
         Tag             =   "9021"
         Top             =   640
         Width           =   4080
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Apólice Cancel/Vencida - Situação "
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   5490
         TabIndex        =   35
         Tag             =   "9024"
         Top             =   915
         Width           =   2805
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Falta assinatura na proposta"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   5490
         TabIndex        =   34
         Tag             =   "9012"
         Top             =   1190
         Width           =   2340
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Alteração da vigência somente com cancelamento da apólice"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   5490
         TabIndex        =   33
         Tag             =   "9029"
         Top             =   1740
         Width           =   4665
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Endereço de cobrança"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   5490
         TabIndex        =   32
         Tag             =   "9036"
         Top             =   1465
         Width           =   2115
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Falta documento original"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   5490
         TabIndex        =   30
         Tag             =   "9037"
         Top             =   90
         Width           =   2610
      End
      Begin VB.ComboBox CmbAssProposta12 
         Height          =   315
         ItemData        =   "T46V147A.frx":0096
         Left            =   8385
         List            =   "T46V147A.frx":00A3
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1125
         Width           =   2220
      End
      Begin VB.TextBox txtPgAntComum9 
         Height          =   270
         Left            =   9315
         MaxLength       =   240
         TabIndex        =   28
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Enviar ao SAC"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   27
         Tag             =   "9041"
         Top             =   75
         Width           =   1410
      End
      Begin VB.ComboBox CmbCNPJ4 
         Height          =   315
         ItemData        =   "T46V147A.frx":00BD
         Left            =   2790
         List            =   "T46V147A.frx":00CD
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   945
         Width           =   2520
      End
      Begin VB.ComboBox CmbEndereco13 
         Height          =   315
         ItemData        =   "T46V147A.frx":00FE
         Left            =   7650
         List            =   "T46V147A.frx":010B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1440
         Width           =   2955
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Faltam documentos da Circular 200"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   2520
         TabIndex        =   24
         Tag             =   "9054"
         Top             =   1737
         Width           =   2820
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "CEP do pernoite do veículo difere do CEP do endereço do segurado"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   135
         TabIndex        =   23
         Tag             =   "9056"
         Top             =   2014
         Width           =   5160
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Nome do proprietário do veículo diverge do nome do segurado"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   17
         Left            =   5490
         TabIndex        =   22
         Tag             =   "9057"
         Top             =   2015
         Width           =   4845
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Existência de mais de dois sinistros nos últimos 2 anos"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   21
         Tag             =   "9058"
         Top             =   2280
         Width           =   5160
      End
      Begin VB.CheckBox ChkComum 
         Caption         =   "Existência de mais de duas restrições no Serasa"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   19
         Left            =   5490
         TabIndex        =   20
         Tag             =   "9059"
         Top             =   2295
         Width           =   3765
      End
   End
   Begin Threed.SSCommand CmdFechar 
      Height          =   375
      Left            =   10110
      TabIndex        =   60
      Top             =   75
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   661
      _StockProps     =   78
      Picture         =   "T46V147A.frx":013D
   End
   Begin Threed.SSCommand CmdSalvar 
      Height          =   375
      Left            =   9555
      TabIndex        =   61
      Top             =   75
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   661
      _StockProps     =   78
      Picture         =   "T46V147A.frx":0603
   End
   Begin VB.Label lblversao 
      AutoSize        =   -1  'True
      Caption         =   "Versão: 05/11/2004"
      Height          =   195
      Left            =   0
      TabIndex        =   214
      Top             =   7830
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Num PI"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   105
      TabIndex        =   213
      Top             =   165
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "Observações"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   2310
      TabIndex        =   212
      Top             =   135
      Width           =   990
   End
End
Attribute VB_Name = "frmT46V147A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim bdSyasEmis      As ADODB.Connection
Dim mLiberaProposta As Boolean
Dim mExistePI       As Boolean
Dim mNumeroPI       As String
Private Sub ChkComum_GotFocus(Index As Integer)
    If ChkComum(Index).Tag = 9054 Then  'Circular 200
        FraComum.Enabled = False
        FraDocto.Visible = True
        If CmdDoctoFec.Visible = True And CmdDoctoFec.Enabled = True Then
            CmdDoctoFec.SetFocus
        End If
    End If
End Sub
Public Sub mpIntanciarForm(pNumeroPI As String)
    mNumeroPI = pNumeroPI
    Me.Show vbModal
End Sub
Private Sub CmdDoctoFec_Click()
    Dim L As Integer

    FraDocto.Tag = ""
    For L = 0 To ChkDocto.Count - 1
        If ChkDocto(L).Value = 1 Then
            FraDocto.Tag = FraDocto.Tag & L & " - " & ChkDocto(L).Caption & ";"
       End If
    Next L
    If FraDocto.Tag = "" Then
        ChkComum(15).Value = 0
    Else
        ChkComum(15).Value = 1
    End If
    FraDocto.Visible = False
    FraComum.Enabled = True
    If ChkComum(6).Visible = True And ChkComum(6).Enabled = True Then
        ChkComum(16).SetFocus
    End If
End Sub
Private Sub CmdFechar_Click()
    Dim QtdeErro    As Integer  'Quantidade de erros encontrados
    Dim I           As Integer  'Utilizado no For...NExt

    Unload Me
End Sub
Function Cdblx(Numero As Variant) As Double
    If IsNumeric(Numero) Then
        Cdblx = Numero
    Else
        Cdblx = 0
    End If
End Function
Function Check(Campo As Field) As Variant
    On Error GoTo TrataCheck
    If IsNull(Campo) Then
        If Campo.Type = 5 Or Campo.Type = 200 Or Campo.Type = 201 Or Campo.Type = 202 Or Campo.Type = 203 Or Campo.Type = 129 Then
            Check = ""
        ElseIf Campo.Type = 7 Then
            Check = ""
        ElseIf Campo.Type = 2 Then
            Check = 0
        ElseIf Campo.Type = 135 Then
        ElseIf Campo.Type = 3 Then
            Check = 0
        Else
            MsgBox "Verificar tipo Campo: " & Campo.Name & " Tipo: " & Campo.Type
            Check = 0
        End If
    Else
        Check = Campo.Value
    End If
    Exit Function
TrataCheck:
    If Err = 3021 Then Resume Next
    MsgBox "Problemas com o campo " & Campo.Name
    Exit Function
End Function
Private Sub CmdSalvar_Click()
    Dim QtdeErro        As Integer
    Dim I               As Integer
    Dim J               As Integer
    Dim NumItem         As Long
    Dim mMotivo         As String
    Dim DataDigitacao   As String
    Dim Aux             As String
    Dim ConsistiuOK     As Boolean
    Dim lMensagem       As String
    Dim lRegistros      As Long

    On Error GoTo CmdSalvar_Err
    If Cdblx(TxtNumPI.Text) = 0 Then
        Call gpGraLog(2, "Favor preencher Número da PI")
        Exit Sub
    End If
        
    For I = 0 To ChkComum.Count - 1
        If ChkComum(I).Value = 1 Then
            QtdeErro = QtdeErro + 1
            Exit For
        End If
    Next I
    For I = 0 To ChkItem.Count - 1
        If ChkItem(I).Value = 1 Then
            QtdeErro = QtdeErro + 1
            Exit For
        End If
    Next I
    For I = 0 To ChkOutrosItems.Count - 1
        If ChkOutrosItems(I).Value = 1 Then
            QtdeErro = QtdeErro + 1
            Exit For
        End If
    Next I
    
    mLiberaProposta = False

    If QtdeErro = 0 Then
        If (Not mExistePI) Then
            Call gpGraLog(2, "Favor escolha um problema na proposta, ou limpe o campo de Observação.")
            Exit Sub
        Else
            Aux = MsgBox("Não há problema selecionado na tela de proposta" & Chr(13) & "Deseja liberar a tela de proposta ?", vbYesNo + vbQuestion)
            If Aux <> 6 Then
                Exit Sub
            Else
                mLiberaProposta = True
                GoTo LiberaProp
            End If
        End If
    End If
    
    For I = 0 To ChkItem.Count - 1
        If ChkItem(I).Value = 1 Then
            If Cdblx(Item1(I)) = 0 And Cdblx(Item2(I)) = 0 And _
                Cdblx(Item3(I)) = 0 And Cdblx(Item4(I)) = 0 And _
                Cdblx(Item5(I)) = 0 And Cdblx(Item6(I)) = 0 And _
                Cdblx(Item7(I)) = 0 And Cdblx(Item8(I)) = 0 And _
                Cdblx(Item9(I)) = 0 And Cdblx(Item10(I)) = 0 Then
                Call gpGraLog(2, "Para problemas no Item, deve ser" & Chr(13) & "colocado o Número Item com problema!")
                If Item1(I).Visible = True And Item1(I).Enabled = True Then
                    Item1(I).SetFocus
                End If
                Exit Sub
            End If
        End If
    Next I

    If ChkComum(4).Value = 1 Then
        If Trim(CmbCNPJ4.Text) = "" Then
            Call gpGraLog(2, "Para o item 'CPF/CNPJ incorreto ou faltante' deve ser selecionado 'na Proposta' ou 'no Perfil'!")
            If CmbCNPJ4.Visible = True And CmbCNPJ4.Enabled = True Then
                CmbCNPJ4.SetFocus
            End If
            Exit Sub
        End If
    End If

    If Trim(CmbCNPJ4.Text) <> "" Then
        If ChkComum(4).Value <> 1 Then
            Call gpGraLog(2, "Clique no item 'CPF/CNPJ incorreto ou faltante' quando selecionar 'na Proposta' ou 'no Perfil'!")
            If ChkComum(4).Visible = True And ChkComum(4).Enabled = True Then
                ChkComum(4).SetFocus
            End If
            Exit Sub
        End If
    End If

    If ChkComum(9).Value = 1 Then
        If Cdblx(txtPgAntComum9.Text) = 0 Then
            Call gpGraLog(2, "Para o item 'Pagto antecipado divergente' deve ser digitado o Valor da Parcela Paga!")
            If txtPgAntComum9.Visible = True And txtPgAntComum9.Enabled = True Then
                txtPgAntComum9.SetFocus
            End If
            Exit Sub
        End If
    End If

    If Cdblx(txtPgAntComum9.Text) <> 0 Then
        If ChkComum(9).Value <> 1 Then
            Call gpGraLog(2, "Clique no item 'Pagto antecipado divergente' se digitar o Valor da Parcela Paga")
            If ChkComum(9).Visible = True And ChkComum(9).Enabled = True Then
                ChkComum(9).SetFocus
            End If
            Exit Sub
        End If
    End If
    If ChkComum(11).Value = 1 Then
        If Trim(txtSituac11.Text) = "" Then
            Call gpGraLog(2, "Para 'Apólice cancelada' deve ser digitado o tipo de cancelamento!")
            If txtSituac11.Visible = True And txtSituac11.Enabled = True Then
                txtSituac11.SetFocus
            End If
            Exit Sub
        End If
    End If

    If Trim(txtSituac11.Text) <> "" Then
        If ChkComum(11).Value <> 1 Then
            Call gpGraLog(2, "Clique 'Apólice cancelada' se colocar situação!")
            If ChkComum(11).Visible = True And ChkComum(11).Enabled = True Then
                ChkComum(11).SetFocus
            End If
            Exit Sub
        End If
    End If
    If ChkComum(12).Value = 1 Then
        If Trim(CmbAssProposta12.Text) = "" Then
            Call gpGraLog(2, "Para o item 'Falta assinatura na proposta' deve selecionar 'Corretor ou Segurado'!")
            If CmbAssProposta12.Visible = True And CmbAssProposta12.Enabled = True Then
                CmbAssProposta12.SetFocus
            End If
            Exit Sub
        End If
    End If

    If Trim(CmbAssProposta12.Text) <> "" Then
        If ChkComum(12).Value <> 1 Then
            Call gpGraLog(2, "Clique no item 'Falta assinatura na proposta' se selecionar 'Corretor ou Segurado'")
            If ChkComum(12).Visible = True And ChkComum(12).Enabled = True Then
                ChkComum(12).SetFocus
            End If
            Exit Sub
        End If
    End If
    If ChkComum(13).Value = 1 Then
        If Trim(CmbEndereco13.Text) = "" Then
            Call gpGraLog(2, "Para o item '" & ChkComum(13).Caption & "' deve selecionar 'Falta endereço completo' ou 'Falta complemento'!")
            If CmbEndereco13.Visible = True And CmbEndereco13.Enabled = True Then
                CmbEndereco13.SetFocus
            End If
            Exit Sub
        End If
    End If
    If Trim(CmbEndereco13.Text) <> "" Then
        If ChkComum(13).Value <> 1 Then
            Call gpGraLog(2, "Clique no item '" & ChkComum(14).Caption & "' se selecionar 'Falta endereço completo' ou 'Falta complemento'")
            If ChkComum(13).Visible = True And ChkComum(13).Enabled = True Then
                ChkComum(13).SetFocus
            End If
            Exit Sub
        End If
    End If
    If ChkComum(1).Value = 1 Then
        If Trim(txtObservacao.Text) = "" Then
            Call gpGraLog(2, "Para o item '" & ChkComum(1).Caption & "' deve ser digitado em 'Observações' " & Chr(13) & "a ação a ser tomada pelo SAC!")
            If txtObservacao.Visible = True And txtObservacao.Enabled = True Then
                txtObservacao.SetFocus
            End If
            Exit Sub
        End If
    End If
'----------
    If ChkOutrosItems(0).Value = 1 Then
        If Cdblx(txtPremio0.Text) = 0 Or Cdblx(txtParc0.Text) = 0 Then
            Call gpGraLog(2, "Para o item 'Cálculo não confere' deve ser digitado o Premio e a Parcela!")
            If txtPremio0.Visible = True And txtPremio0.Enabled = True Then
                txtPremio0.SetFocus
            End If
            Exit Sub
        End If
    End If

    If Cdblx(txtPremio0.Text) <> 0 Or Cdblx(txtParc0.Text) <> 0 Then
        If ChkOutrosItems(0).Value <> 1 Then
            Call gpGraLog(2, "Clique no item 'Cálculo não confere' se digitar o Premio e a Parcela!")
            If ChkComum(0).Visible = True And ChkComum(0).Enabled = True Then
                ChkComum(0).SetFocus
            End If
            Exit Sub
        End If
    End If
    If ChkOutrosItems(1).Value = 1 Then
        If Cdblx(txtParc1.Text) = 0 Or Cdblx(txtValor1.Text) = 0 Then
            Call gpGraLog(2, "Para 'Parcelamento x Vigência' deve ser digitado as Parcelas e o Valor!")
            If txtParc1.Visible = True And txtParc1.Enabled = True Then
                txtParc1.SetFocus
            End If
            Exit Sub
        End If
    End If
    
    If Cdblx(txtParc1.Text) <> 0 Or Cdblx(txtValor1.Text) <> 0 Then
        If ChkOutrosItems(1).Value <> 1 Then
            Call gpGraLog(2, "Clique 'Parcelamento x Vigência' se digitar as Parcelas e o Valor!")
            If ChkOutrosItems(1).Visible = True And ChkOutrosItems(1).Enabled = True Then
                ChkOutrosItems(1).SetFocus
            End If
            Exit Sub
        End If
    End If

    If ChkOutrosItems(2).Value = 1 Then
        If Cdblx(txtPgAntItem2.Text) = 0 Then
            Call gpGraLog(2, "Para o item 'Pagto antecipado divergente' deve ser digitado o Valor da Parcela Paga!")
            If txtPgAntItem2.Visible = True And txtPgAntItem2.Enabled = True Then
                txtPgAntItem2.SetFocus
            End If
            Exit Sub
        End If
    End If

    If Cdblx(txtPgAntItem2.Text) <> 0 Then
        If ChkOutrosItems(2).Value <> 1 Then
            Call gpGraLog(2, "Clique no item 'Pagto antecipado divergente' se digitar o Valor da Parcela Paga")
            If ChkOutrosItems(2).Visible = True And ChkOutrosItems(2).Enabled = True Then
                ChkOutrosItems(2).SetFocus
            End If
            Exit Sub
        End If
    End If
    
    If ChkOutrosItems(3).Value = 1 Then
        If Trim(CmbApolice3.Text) = "" Then
            Call gpGraLog(2, "Para o item '" & ChkOutrosItems(3).Caption & "' deve selecionar 'Falta' ou 'Ilegível'!")
            If CmbApolice3.Visible = True And CmbApolice3.Enabled = True Then
                CmbApolice3.SetFocus
            End If
            Exit Sub
        End If
    End If
    If Trim(CmbApolice3.Text) <> "" Then
        If ChkOutrosItems(3).Value <> 1 Then
            Call gpGraLog(2, "Clique no item '" & ChkOutrosItems(3).Caption & "' se selecionar 'Falta' ou 'Ilegível'")
            If ChkOutrosItems(3).Visible = True And ChkOutrosItems(3).Enabled = True Then
                ChkOutrosItems(3).SetFocus
            End If
            Exit Sub
        End If
    End If
    If ChkOutrosItems(5).Value = 1 Then
        If Trim(txtObservacao.Text) = "" Then
            Call gpGraLog(2, "Para o item '" & ChkOutrosItems(5).Caption & "' deve ser digitado em 'Observações' " & Chr(13) & "a ação a ser tomada pelo SAC!")
            If txtObservacao.Visible = True And txtObservacao.Enabled = True Then
                txtObservacao.SetFocus
            End If
            Exit Sub
        End If
    End If
'----------
    If ChkItem(3).Value = 1 Then
        If Trim(txtDados3.Text) = "" Then
            Call gpGraLog(2, "Para 'Dados Divegentes' deve ser digitado as informações divergentes do perfil!")
            If txtDados3.Visible = True And txtDados3.Enabled = True Then
                txtDados3.SetFocus
            End If
            Exit Sub
        End If
    End If
    If Trim(txtDados3.Text) <> "" Then
       If ChkItem(3).Value <> 1 Then
            Call gpGraLog(2, "Clique 'Dados Divegentes' quando digitar as informações divergentes do perfil!")
            If ChkItem(3).Visible = True And ChkItem(3).Enabled = True Then
                ChkItem(3).SetFocus
            End If
            Exit Sub
        End If
    End If
    If ChkItem(4).Value = 1 Then
        If Trim(CmbAssPerfil4.Text) = "" Then
            Call gpGraLog(2, "Para o item 'Perfil do Segurado - Assinaturas Faltantes' deve selecionar 'Corretor ou Segurado'!")
            If CmbAssPerfil4.Visible = True And CmbAssPerfil4.Enabled = True Then
                CmbAssPerfil4.SetFocus
            End If
            Exit Sub
        End If
    End If
    If Trim(CmbAssPerfil4.Text) <> "" Then
        If ChkItem(4).Value <> 1 Then
            Call gpGraLog(2, "Clique no item 'Perfil do Segurado - Assinaturas Faltantes' se selecionar 'Corretor ou Segurado'")
            If ChkItem(4).Visible = True And ChkItem(4).Enabled = True Then
                ChkItem(4).SetFocus
            End If
            Exit Sub
        End If
    End If
    If ChkItem(7).Value = 1 Then
        If Cdblx(txtApolice7.Text) = 0 Then
            Call gpGraLog(2, "Para 'Placa e/ou Chassi - duplicidade' deve ser digitado a Apólice!")
            If txtApolice7.Visible = True And txtApolice7.Enabled = True Then
                txtApolice7.SetFocus
            End If
            Exit Sub
        End If
    End If
    If Cdblx(txtApolice7.Text) <> 0 Then
       If ChkItem(7).Value <> 1 Then
            Call gpGraLog(2, "Clique 'Placa e/ou Chassi - duplicidade' quando digitar a Apólice!")
            If ChkItem(7).Visible = True And ChkItem(7).Enabled = True Then
                ChkItem(7).SetFocus
            End If
            Exit Sub
        End If
    End If
   
    If ChkItem(8).Value = 1 Then
        If Trim(CmbTipInformar8.Text) = "" Then
            Call gpGraLog(2, "Para 'Placa e/ou Chassi - Informar' deve ser selecionado Placa ou Chassi!")
            If CmbTipInformar8.Visible = True And CmbTipInformar8.Enabled = True Then
                CmbTipInformar8.SetFocus
            End If
            Exit Sub
        End If
    End If
    If Trim(CmbTipInformar8.Text) <> "" Then
       If ChkItem(8).Value <> 1 Then
            Call gpGraLog(2, "Clique 'Placa e/ou Chassi - Informar' quando selecionar Placa ou Chassi!")
            If ChkItem(8).Visible = True And ChkItem(8).Enabled = True Then
                ChkItem(8).SetFocus
            End If
            Exit Sub
        End If
    End If
   
    DataDigitacao = Format(Date, "YYYYMMDD")
    
LiberaProp:
    
    'bdSYASEMIS.BeginTrans
    
    gDelete = "Delete from ged.dbo.Tab_Prop_Mensagem  Where num_pi = " & TxtNumPI.Text & " and "
    gDelete = gDelete & " cod_mensagem > 9000  and   cod_mensagem <> 9500  and  cod_tip_mensagem = 2 and "
    gDelete = gDelete & " dat_liber = 0  "
    
    ''----------------------------------------------------------------------------
    If gfExeSQL(bdSyasEmis, gDelete, lMensagem, lRegistros) <> 0 Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If
    ''----------------------------------------------------------------------------

    gDelete = "Delete from ged.dbo.Tab_Prop_Mensagem  Where num_pi = " & TxtNumPI.Text & " and "
    gDelete = gDelete & " cod_mensagem = 9111 "
    gDelete = gDelete & " and dat_liber = 0 "
    
    ''----------------------------------------------------------------------------
    If gfExeSQL(bdSyasEmis, gDelete, lMensagem, lRegistros) <> 0 Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If
    ''----------------------------------------------------------------------------
    
    If mLiberaProposta Then
        ConsistiuOK = False  'para dar o enter no final
        DoEvents
        Call gpFecha2(bdSyasEmis)
        End
    End If
    
    If Trim(txtObservacao.Text) <> "" Then
        mMotivo = Replace(txtObservacao.Text, "'", "´")
        gInsert = "Insert into ged.dbo.Tab_Prop_Mensagem  values ("
        gInsert = gInsert & TxtNumPI.Text & ", 0000,  9111 , 2 , '" & mMotivo & "', " & DataDigitacao & ", 0 ,0 )"
        ''----------------------------------------------------------------------------
        If gfExeSQL(bdSyasEmis, gInsert, lMensagem, lRegistros) <> 0 Then
            Call gpGraLog(2, lMensagem)
            Exit Sub
        End If
        ''----------------------------------------------------------------------------
    End If
    
    
    For I = 0 To ChkComum.Count - 1
        mMotivo = ""
        If ChkComum(I).Value = 1 Then
            If I = 4 Then
                mMotivo = CmbCNPJ4.Text
            End If
            If I = 9 Then
                mMotivo = "Valor da Parcela: " & txtPgAntComum9.Text
            End If
            If I = 11 Then
                mMotivo = "Situação: " & txtSituac11.Text
            End If
            If I = 12 Then
                mMotivo = "Falta Assinatura: " & CmbAssProposta12.Text
            End If
            If I = 13 Then
                mMotivo = CmbEndereco13.Text
            End If
            If I = 15 Then
                mMotivo = Mid(FraDocto.Tag, 1, 240)
            End If
            gInsert = "Insert into ged.dbo.Tab_Prop_Mensagem  values ("
            gInsert = gInsert & TxtNumPI.Text & ", 0000, " & ChkComum(I).Tag & ", 2 , '" & mMotivo & "', " & DataDigitacao & ", 0 ,0 )"
            
            ''----------------------------------------------------------------------------
            If gfExeSQL(bdSyasEmis, gInsert, lMensagem, lRegistros) <> 0 Then
                Call gpGraLog(2, lMensagem)
                Exit Sub
            End If
            ''----------------------------------------------------------------------------
        End If
    Next I
    For I = 0 To ChkOutrosItems.Count - 1
        mMotivo = ""
        If ChkOutrosItems(I).Value = 1 Then
            If I = 0 Then
                mMotivo = "Premio: " & txtPremio0.Text & " - Parcelas: " & txtParc0.Text
            End If
            If I = 1 Then
                mMotivo = "Parcelas: " & txtParc1.Text & " - Valor: " & txtValor1.Text
            End If
            If I = 2 Then
                mMotivo = "Valor da Parcela: " & txtPgAntItem2.Text
            End If
            If I = 3 Then
                mMotivo = CmbApolice3.Text
            End If
            gInsert = "Insert into ged.dbo.Tab_Prop_Mensagem  values ("
            gInsert = gInsert & TxtNumPI.Text & ", 9999, " & ChkOutrosItems(I).Tag & ", 2 , '" & mMotivo & "', " & DataDigitacao & ", 0 , 0 )"
            ''----------------------------------------------------------------------------
            If gfExeSQL(bdSyasEmis, gInsert, lMensagem, lRegistros) <> 0 Then
                Call gpGraLog(2, lMensagem)
                Exit Sub
            End If
            ''----------------------------------------------------------------------------
        End If
    Next I
'-----
    For I = 0 To ChkItem.Count - 1
        If ChkItem(I).Value = 1 Then
            For J = 1 To 10
                mMotivo = ""
                NumItem = 0
                If J = 1 Then NumItem = Cdblx(Item1(I))
                If J = 2 Then NumItem = Cdblx(Item2(I))
                If J = 3 Then NumItem = Cdblx(Item3(I))
                If J = 4 Then NumItem = Cdblx(Item4(I))
                If J = 5 Then NumItem = Cdblx(Item5(I))
                If J = 6 Then NumItem = Cdblx(Item6(I))
                If J = 7 Then NumItem = Cdblx(Item7(I))
                If J = 8 Then NumItem = Cdblx(Item8(I))
                If J = 9 Then NumItem = Cdblx(Item9(I))
                If J = 10 Then NumItem = Cdblx(Item10(I))
                If NumItem > 0 Then
                    If I = 3 Then
                        mMotivo = "Dados: " & txtDados3.Text
                    End If
                    If I = 4 Then
                        mMotivo = "Perfil não assinado pelo: " & CmbAssPerfil4
                    End If
                    If I = 7 Then
                        mMotivo = "Apólice com duplicidade: " & txtApolice7
                    End If
                    If I = 8 Then
                        mMotivo = CmbTipInformar8.Text
                    End If
                    If I = 14 Then
                        mMotivo = "Falta Assinatura: " & CmbAssProposta12.Text
                    End If
                    
                    gInsert = "Insert into ged.dbo.Tab_Prop_Mensagem  values ("
                    gInsert = gInsert & TxtNumPI.Text & ", " & NumItem & " , " & ChkItem(I).Tag & ",2, '" & mMotivo & "', " & DataDigitacao & ", 0, 0  )"
                    
                    ''----------------------------------------------------------------------------
                    If gfExeSQL(bdSyasEmis, gInsert, lMensagem, lRegistros) <> 0 Then
                        Call gpGraLog(2, lMensagem)
                        Exit Sub
                    End If
                    ''----------------------------------------------------------------------------
                End If
            Next
        End If
    Next I
    ConsistiuOK = True
    DoEvents
    Unload Me
    Exit Sub

CmdSalvar_Err:
    If Err <> 0 Then
        MsgBox Err & " - " & Error & Chr(13) & "Erro ao salvar problemas com proposta - " & gInsert
        bdSyasEmis.RollbackTrans
        Resume Next
    End If
End Sub
Public Sub ConfiguracoesLocais()
    Dim lData1          As String   'Auxiliar para verificar configuração de data.
    Dim lData2          As String   'Auxiliar para verificar configuração de data.
    Dim lValor          As String   'Auxiliar para verificar configuração de número.
       
    '5. Verifica configuração de data: dd/mm/aaaa.
    If Len(Date) <> 10 Then
        Call gpGraLog(4, "Microcomputador não configurado para datas 'dd/mm/aaaa' (10 posições). " & _
                         "Entre em >>>Configurações >>>Painel de controle >>>Configurações " & _
                         "regionais >>>Data e altere data para 'dd/mm/aaaa'.")  'Encerra programa.
    End If
    If lData1 <> lData2 Then
        Call gpGraLog(4, "Microcomputador não configurado para datas 'dd/mm/aaaa' (ordem do " & _
                         "dia/mês/ano). Entre em >>>Configurações >>>Painel de controle " & _
                         ">>>Configurações regionais >>>Data e altere data para 'dd/mm/aaaa'.")
        'Encerra programa.
    End If

    '6. Verifica configuração de números: ponto como separador de milhar e vírgula como separador de
    '   decimal.
    lValor = Trim$(Format$(1000, "standard"))
    If Mid$(lValor, 2, 1) <> "." Then
        Call gpGraLog(4, "Microcomputador não configurado para número com separador de milhar igual " & _
                         "a ponto. Entre em >>>Configurações >>>Painel de controle >>>Configurações " & _
                         "regionais >>>Número e altere 'símbolo de agrupamento de dígitos' para ponto.")
        'Encerra programa.
    End If
    If Mid$(lValor, 6, 1) <> "," Then
        Call gpGraLog(4, "Microcomputador não configurado para número com separador de decimal " & _
                         "igual a vírgula. Entre em >>>Configurações >>>Painel de controle " & _
                         ">>>Configurações regionais >>>Número e altere 'símbolo decimal' para vírgula.")
        'Encerra programa.
    End If

   
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
End Sub
Private Sub Form_Load()
    Dim lMensagem   As String
    Dim sTitle      As String

    On Error GoTo Form_Load_Err
'    If App.PrevInstance = True Then
'        MsgBox "Aplicativo já está ativo "
'        sTitle = Me.Caption
'        Me.Caption = Hex$(Me.hwnd)
'        AppActivate sTitle
'        End
'    End If

    Call gpCenForm(Me)

    mExistePI = False
    mLiberaProposta = False
    ConfiguracoesLocais
    Screen.MousePointer = vbHourglass
    If gfAbrBasDados("SYAS_EMIS", gTipSYASEMIS, gEndSYASEMIS, bdSyasEmis, lMensagem) = False Then
        Call gpGraLog(1, lMensagem)
        Exit Sub
    End If
    FraDocto.Visible = False
    Call CarregarForm(mNumeroPI)
    Screen.MousePointer = vbDefault
    Exit Sub

Form_Load_Err:
    MsgBox Err & ":" & Error & " - " & " Erro carregando Sistema"
    End
End Sub
Function StrZero(Numero As String, QtdDecimal As Integer) As String
    If Cdblx(Numero) = 0 Then Numero = 0
    StrZero = Format(Numero, String(QtdDecimal, "0"))
End Function

Private Sub CarregarForm(pNumeroPI As String)
On Error GoTo ProcuraParametros_Err

Dim I, L  As Integer
Dim NumItem  As String
Dim AuxPI As String
Dim NumErr As Integer
Dim mMotivo As String
Dim Pos, Pos2 As Integer
Dim lMensagem As String
    
    TxtNumPI.Text = pNumeroPI
   
    Err = 0
    Dim RstProp    As New ADODB.Recordset
    Set RstProp = New ADODB.Recordset
   
    
    gSelect = "Select * from ged.dbo.Tab_Prop_Mensagem  Where num_pi = " & TxtNumPI.Text
    gSelect = gSelect & " and cod_mensagem > 9000   "
    gSelect = gSelect & " and dat_liber = 0  "
    gSelect = gSelect & " order by item_pi"
    NumErr = 7
    
    If gfObtRegistro(bdSyasEmis, gSelect, RstProp, lMensagem) = False Then
        Call gpFecha2(bdSyasEmis)
        Call gpMensagem(lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    NumErr = 8
    Do While Not RstProp.EOF
        mExistePI = True
        If RstProp!Item_Pi = 0 Then
            If RstProp!Cod_Mensagem = 9111 Then
                txtObservacao = Check(RstProp!Motivo)
            Else
                For I = 0 To ChkComum.Count - 1
                    If Cdblx(ChkComum(I).Tag) = RstProp!Cod_Mensagem Then
                        ChkComum(I).Value = 1
                        mMotivo = Check(RstProp!Motivo)
                        If I = 4 And Trim(mMotivo) <> "" Then
                            If Trim(mMotivo) = "Ambos" Then
                                CmbCNPJ4.ListIndex = 1
                            ElseIf Trim(mMotivo) = "na Proposta" Then
                                CmbCNPJ4.ListIndex = 2
                            ElseIf Trim(mMotivo) = "no Perfil do Segurado" Then
                                CmbCNPJ4.ListIndex = 3
                            End If
                        End If
                        If I = 9 And Trim(mMotivo) <> "" Then
                            Pos = InStr(mMotivo, ":") + 1
                            txtPgAntComum9.Text = Trim(Mid(mMotivo, Pos))
                        End If
                        If I = 11 And Trim(mMotivo) <> "" Then
                            Pos = InStr(mMotivo, ":") + 1
                            txtSituac11.Text = Trim(Mid(mMotivo, Pos))
                        End If
                        If I = 12 And Trim(mMotivo) <> "" Then
                            Pos = InStr(mMotivo, ":") + 1
                            If Trim(Mid(mMotivo, Pos)) = "Segurado" Then
                                CmbAssProposta12.ListIndex = 2
                            ElseIf Trim(Mid(mMotivo, Pos)) = "Corretor" Then
                                CmbAssProposta12.ListIndex = 1
                            End If
                        End If
                        If I = 13 And Trim(mMotivo) <> "" Then
                            If Trim(mMotivo) = "Falta complemento" Then
                                CmbEndereco13.ListIndex = 1
                            ElseIf Trim(mMotivo) = "Falta endereço completo" Then
                                CmbEndereco13.ListIndex = 2
                            End If
                        End If
                        If I = 15 And Trim(mMotivo) <> "" Then
                             For L = 0 To ChkDocto.Count - 1
                                 If Trim(mMotivo) <> "" Then
                                    Pos = InStr(mMotivo, "-") - 1
                                    If Pos <= 0 Then Pos = 2
                                    If Mid(mMotivo, 1, Pos) = L Then
                                       ChkDocto(L).Value = 1
                                       If InStr(mMotivo, ";") <> 0 Then
                                            mMotivo = Mid(mMotivo, InStr(mMotivo, ";") + 1)
                                       Else
                                           mMotivo = " "
                                       End If
                                    End If
                                End If
                             Next L
                        End If
                        Exit For
                    End If
                Next I
            End If
        ElseIf RstProp!Item_Pi = 9999 Then
            For I = 0 To ChkOutrosItems.Count - 1
                If Cdblx(ChkOutrosItems(I).Tag) = RstProp!Cod_Mensagem Then
                    ChkOutrosItems(I).Value = 1
                    mMotivo = Check(RstProp!Motivo)
                    If I = 0 And Trim(mMotivo) <> "" Then
                        Pos = InStr(mMotivo, ":") + 1
                        Pos2 = InStr(mMotivo, "-") - Pos
                        txtPremio0.Text = Trim(Mid(mMotivo, Pos, Pos2))
                        Pos = InStrRev(mMotivo, ":") + 1
                        txtParc0.Text = Trim(Mid(mMotivo, Pos))
                    End If
                    If I = 1 And Trim(mMotivo) <> "" Then
                        Pos = InStr(mMotivo, ":") + 1
                        Pos2 = InStr(mMotivo, "-") - Pos
                        txtParc1.Text = Trim(Mid(mMotivo, Pos, Pos2))
                        Pos = InStrRev(mMotivo, ":") + 1
                        txtValor1.Text = Trim(Mid(mMotivo, Pos))
                    End If
                    If I = 2 And Trim(mMotivo) <> "" Then
                        Pos = InStr(mMotivo, ":") + 1
                        txtPgAntItem2.Text = Trim(Mid(mMotivo, Pos))
                    End If
                    If I = 3 And Trim(mMotivo) <> "" Then
                        If Trim(mMotivo) = "Falta" Then
                            CmbApolice3.ListIndex = 1
                        ElseIf Trim(mMotivo) = "Ilegível" Then
                            CmbApolice3.ListIndex = 2
                        End If
                    End If
                    Exit For
                End If
            Next I
        Else
            For I = 0 To ChkItem.Count - 1
                If Cdblx(ChkItem(I).Tag) = RstProp!Cod_Mensagem Then
                    ChkItem(I).Value = 1
                    mMotivo = Check(RstProp!Motivo)
                    If I = 3 And Trim(mMotivo) <> "" Then
                        Pos = InStr(mMotivo, ":") + 1
                        txtDados3.Text = Trim(Mid(mMotivo, Pos))
                    End If
                    If I = 4 And Trim(mMotivo) <> "" Then
                        Pos = InStr(mMotivo, ":") + 1
                        If Trim(Mid(mMotivo, Pos)) = "Segurado" Then
                            CmbAssPerfil4.ListIndex = 3
                        ElseIf Trim(Mid(mMotivo, Pos)) = "Corretor" Then
                            CmbAssPerfil4.ListIndex = 2
                        ElseIf Trim(Mid(mMotivo, Pos)) = "Ambos" Then
                            CmbAssPerfil4.ListIndex = 1
                        End If
                    End If
                    If I = 7 And Trim(mMotivo) <> "" Then
                        Pos = InStr(mMotivo, ":") + 1
                        txtApolice7.Text = Trim(Mid(mMotivo, Pos))
                    End If
                    If I = 8 And Trim(mMotivo) <> "" Then
                        If Trim(mMotivo) = "Chassi" Then CmbTipInformar8.ListIndex = 1
                        If Trim(mMotivo) = "Placa" Then CmbTipInformar8.ListIndex = 3
                        If Trim(mMotivo) = "Chassi e Placa" Then CmbTipInformar8.ListIndex = 2
                    End If
                    
                    NumItem = StrZero(RstProp!Item_Pi, 4)
                    If Cdblx(Item1(I).Text) = 0 Then
                        Item1(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item2(I).Text) = 0 Then
                        Item2(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item3(I).Text) = 0 Then
                        Item3(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item4(I).Text) = 0 Then
                        Item4(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item5(I).Text) = 0 Then
                        Item5(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item6(I).Text) = 0 Then
                        Item6(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item7(I).Text) = 0 Then
                        Item7(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item8(I).Text) = 0 Then
                        Item8(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item9(I).Text) = 0 Then
                        Item9(I).Text = NumItem
                        Exit For
                    End If
                    If Cdblx(Item10(I).Text) = 0 Then
                        Item10(I).Text = NumItem
                        Exit For
                    End If
                    Exit For
                End If
            Next I
        End If
        RstProp.MoveNext
    Loop
    NumErr = 5
    
    RstProp.Close
    Set RstProp = Nothing
    
    Exit Sub
    
ProcuraParametros_Err:
If Err = 430 Then Exit Sub
MsgBox Err & ": " & Error & " - Procura Parametros  " & NumErr
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim QtdeErro As Integer
Dim I           As Integer

    'For I = 0 To ChkComum.Count - 1
    '    If ChkComum(I).Value = 1 Then
    '        QtdeErro = QtdeErro + 1
    '        Exit For
    '    End If
    'Next I'

    'For I = 0 To ChkItem.Count - 1
       ' If ChkItem(I).Value = 1 Then
     '       QtdeErro = QtdeErro + 1
      '      Exit For
        'End If
'    Next I
 '   For I = 0 To ChkOutrosItems.Count - 1
  '      If ChkOutrosItems(I).Value = 1 Then
   '         QtdeErro = QtdeErro + 1
    '        Exit For
     '   End If
    'Next I

    
If QtdeErro <> 0 Then
    MsgBox "ATENÇÃO " & vbCrLf & vbCrLf & "Existem itens selecionados na tela. Caso algum desses itens ainda não tenha sido gravado, o mesmo será DESCONSIDERADO!", vbInformation + vbOKOnly, "Informação"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo TratarError
          
        
   Call gpFecha2(bdSyasEmis)
   
Exit Sub
TratarError:
  MsgBox "Erro: " & Err.Description & ")", vbExclamation
  
End Sub
Function ConsisteKeyNum(Tecla As Integer) As Integer
    If (Tecla < 48 Or Tecla > 57) And Tecla <> 9 And Tecla <> 8 And Tecla <> 44 Then
        ConsisteKeyNum = 0
    Else
        ConsisteKeyNum = Tecla
    End If
End Function
Private Sub Item1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item1_LostFocus(Index As Integer)
    If Cdblx(Item1(Index)) > 0 Then
        Item1(Index) = StrZero(Item1(Index), 4)
    End If
End Sub
Private Sub Item2_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item2_LostFocus(Index As Integer)
    If Cdblx(Item2(Index)) > 0 Then
        Item2(Index) = StrZero(Item2(Index), 4)
    End If
End Sub
Private Sub Item3_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item3_LostFocus(Index As Integer)
    If Cdblx(Item3(Index)) > 0 Then
        Item3(Index) = StrZero(Item3(Index), 4)
    End If
End Sub
Private Sub Item4_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item4_LostFocus(Index As Integer)
    If Cdblx(Item4(Index)) > 0 Then
        Item4(Index) = StrZero(Item4(Index), 4)
    End If
End Sub
Private Sub Item5_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item5_LostFocus(Index As Integer)
    If Cdblx(Item5(Index)) > 0 Then
        Item5(Index) = StrZero(Item5(Index), 4)
    End If
End Sub
Private Sub Item6_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item6_LostFocus(Index As Integer)
    If Cdblx(Item6(Index)) > 0 Then
        Item6(Index) = StrZero(Item6(Index), 4)
    End If
End Sub
Private Sub Item7_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item7_LostFocus(Index As Integer)
    If Cdblx(Item7(Index)) > 0 Then
        Item7(Index) = StrZero(Item7(Index), 4)
    End If
End Sub
Private Sub Item8_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item8_LostFocus(Index As Integer)
    If Cdblx(Item8(Index)) > 0 Then
        Item8(Index) = StrZero(Item8(Index), 4)
    End If
End Sub
Private Sub Item9_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item9_LostFocus(Index As Integer)
    If Cdblx(Item9(Index)) > 0 Then
        Item9(Index) = StrZero(Item9(Index), 4)
    End If
End Sub
Private Sub Item10_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub Item10_LostFocus(Index As Integer)
    If Cdblx(Item10(Index)) > 0 Then
        Item10(Index) = StrZero(Item10(Index), 4)
    End If
End Sub

Private Sub txtApolice7_KeyPress(KeyAscii As Integer)
   KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub

Private Sub txtParc1_KeyPress(KeyAscii As Integer)
   KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub
Private Sub txtPgAntComum9_KeyPress(KeyAscii As Integer)
   KeyAscii = ConsisteKeyNum(KeyAscii)
End Sub


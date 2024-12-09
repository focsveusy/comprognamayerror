<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Books
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Books))
        Me.dgBooksList = New System.Windows.Forms.DataGridView()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtPrice = New System.Windows.Forms.TextBox()
        Me.cboCategories = New System.Windows.Forms.ComboBox()
        Me.txtAuthor = New System.Windows.Forms.TextBox()
        Me.txtBTitle = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtQuantity = New System.Windows.Forms.TextBox()
        Me.btnBRefresh = New System.Windows.Forms.Button()
        Me.btnBSave = New System.Windows.Forms.Button()
        Me.btnBEdit = New System.Windows.Forms.Button()
        Me.btnBDelete = New System.Windows.Forms.Button()
        Me.btnBAdd = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.btnBDashboard = New System.Windows.Forms.Button()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.btnBUsers = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.btnBBooks = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnclear = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtBSearch = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.btnBBookshop = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnBLogout = New System.Windows.Forms.Button()
        CType(Me.dgBooksList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel7.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel8.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgBooksList
        '
        Me.dgBooksList.BackgroundColor = System.Drawing.Color.White
        Me.dgBooksList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgBooksList.Location = New System.Drawing.Point(50, 239)
        Me.dgBooksList.Margin = New System.Windows.Forms.Padding(2)
        Me.dgBooksList.Name = "dgBooksList"
        Me.dgBooksList.RowHeadersWidth = 51
        Me.dgBooksList.RowTemplate.Height = 24
        Me.dgBooksList.Size = New System.Drawing.Size(582, 187)
        Me.dgBooksList.TabIndex = 20
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(301, 174)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(96, 23)
        Me.Label11.TabIndex = 18
        Me.Label11.Text = "Books List"
        '
        'txtPrice
        '
        Me.txtPrice.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPrice.Font = New System.Drawing.Font("Century Gothic", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrice.Location = New System.Drawing.Point(569, 98)
        Me.txtPrice.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.Size = New System.Drawing.Size(81, 25)
        Me.txtPrice.TabIndex = 10
        '
        'cboCategories
        '
        Me.cboCategories.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCategories.Font = New System.Drawing.Font("Century Gothic", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCategories.FormattingEnabled = True
        Me.cboCategories.Items.AddRange(New Object() {"English", "Mathematics", "Programming", "Statistics", "Science", "Novels"})
        Me.cboCategories.Location = New System.Drawing.Point(333, 97)
        Me.cboCategories.Margin = New System.Windows.Forms.Padding(2)
        Me.cboCategories.Name = "cboCategories"
        Me.cboCategories.Size = New System.Drawing.Size(133, 25)
        Me.cboCategories.TabIndex = 8
        '
        'txtAuthor
        '
        Me.txtAuthor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAuthor.Font = New System.Drawing.Font("Century Gothic", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAuthor.Location = New System.Drawing.Point(202, 97)
        Me.txtAuthor.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAuthor.Name = "txtAuthor"
        Me.txtAuthor.Size = New System.Drawing.Size(120, 25)
        Me.txtAuthor.TabIndex = 7
        '
        'txtBTitle
        '
        Me.txtBTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBTitle.Font = New System.Drawing.Font("Century Gothic", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBTitle.Location = New System.Drawing.Point(74, 97)
        Me.txtBTitle.Margin = New System.Windows.Forms.Padding(2)
        Me.txtBTitle.Name = "txtBTitle"
        Me.txtBTitle.Size = New System.Drawing.Size(115, 25)
        Me.txtBTitle.TabIndex = 6
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Century Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(567, 77)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(42, 19)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "Price"
        '
        'txtQuantity
        '
        Me.txtQuantity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtQuantity.Font = New System.Drawing.Font("Century Gothic", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuantity.Location = New System.Drawing.Point(479, 98)
        Me.txtQuantity.Margin = New System.Windows.Forms.Padding(2)
        Me.txtQuantity.Name = "txtQuantity"
        Me.txtQuantity.Size = New System.Drawing.Size(76, 25)
        Me.txtQuantity.TabIndex = 9
        '
        'btnBRefresh
        '
        Me.btnBRefresh.BackColor = System.Drawing.Color.FromArgb(CType(CType(104, Byte), Integer), CType(CType(140, Byte), Integer), CType(CType(202, Byte), Integer))
        Me.btnBRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBRefresh.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBRefresh.ForeColor = System.Drawing.Color.White
        Me.btnBRefresh.Location = New System.Drawing.Point(451, 211)
        Me.btnBRefresh.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBRefresh.Name = "btnBRefresh"
        Me.btnBRefresh.Size = New System.Drawing.Size(78, 23)
        Me.btnBRefresh.TabIndex = 17
        Me.btnBRefresh.Text = "Refresh"
        Me.btnBRefresh.UseVisualStyleBackColor = False
        '
        'btnBSave
        '
        Me.btnBSave.BackColor = System.Drawing.Color.FromArgb(CType(CType(104, Byte), Integer), CType(CType(140, Byte), Integer), CType(CType(202, Byte), Integer))
        Me.btnBSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBSave.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBSave.ForeColor = System.Drawing.Color.White
        Me.btnBSave.Location = New System.Drawing.Point(122, 141)
        Me.btnBSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBSave.Name = "btnBSave"
        Me.btnBSave.Size = New System.Drawing.Size(85, 23)
        Me.btnBSave.TabIndex = 16
        Me.btnBSave.Text = "Save"
        Me.btnBSave.UseVisualStyleBackColor = False
        '
        'btnBEdit
        '
        Me.btnBEdit.BackColor = System.Drawing.Color.FromArgb(CType(CType(104, Byte), Integer), CType(CType(140, Byte), Integer), CType(CType(202, Byte), Integer))
        Me.btnBEdit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBEdit.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBEdit.ForeColor = System.Drawing.Color.White
        Me.btnBEdit.Location = New System.Drawing.Point(221, 141)
        Me.btnBEdit.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBEdit.Name = "btnBEdit"
        Me.btnBEdit.Size = New System.Drawing.Size(78, 23)
        Me.btnBEdit.TabIndex = 15
        Me.btnBEdit.Text = "Edit"
        Me.btnBEdit.UseVisualStyleBackColor = False
        '
        'btnBDelete
        '
        Me.btnBDelete.BackColor = System.Drawing.Color.FromArgb(CType(CType(104, Byte), Integer), CType(CType(140, Byte), Integer), CType(CType(202, Byte), Integer))
        Me.btnBDelete.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(104, Byte), Integer), CType(CType(140, Byte), Integer), CType(CType(202, Byte), Integer))
        Me.btnBDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBDelete.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBDelete.ForeColor = System.Drawing.Color.White
        Me.btnBDelete.Location = New System.Drawing.Point(409, 141)
        Me.btnBDelete.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBDelete.Name = "btnBDelete"
        Me.btnBDelete.Size = New System.Drawing.Size(76, 23)
        Me.btnBDelete.TabIndex = 14
        Me.btnBDelete.Text = "Delete"
        Me.btnBDelete.UseVisualStyleBackColor = False
        '
        'btnBAdd
        '
        Me.btnBAdd.BackColor = System.Drawing.Color.FromArgb(CType(CType(104, Byte), Integer), CType(CType(140, Byte), Integer), CType(CType(202, Byte), Integer))
        Me.btnBAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBAdd.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBAdd.ForeColor = System.Drawing.Color.White
        Me.btnBAdd.Location = New System.Drawing.Point(312, 141)
        Me.btnBAdd.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBAdd.Name = "btnBAdd"
        Me.btnBAdd.Size = New System.Drawing.Size(85, 23)
        Me.btnBAdd.TabIndex = 13
        Me.btnBAdd.Text = "Add"
        Me.btnBAdd.UseVisualStyleBackColor = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Century Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(475, 77)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(70, 19)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "Quantity"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Century Gothic", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(240, 18)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(150, 18)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "BookShop Software"
        '
        'Panel7
        '
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel7.Controls.Add(Me.btnBDashboard)
        Me.Panel7.Location = New System.Drawing.Point(2, 277)
        Me.Panel7.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(144, 42)
        Me.Panel7.TabIndex = 7
        '
        'btnBDashboard
        '
        Me.btnBDashboard.BackColor = System.Drawing.Color.FromArgb(CType(CType(63, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(126, Byte), Integer))
        Me.btnBDashboard.FlatAppearance.BorderSize = 0
        Me.btnBDashboard.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBDashboard.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBDashboard.ForeColor = System.Drawing.Color.White
        Me.btnBDashboard.Location = New System.Drawing.Point(-1, 2)
        Me.btnBDashboard.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBDashboard.Name = "btnBDashboard"
        Me.btnBDashboard.Size = New System.Drawing.Size(146, 39)
        Me.btnBDashboard.TabIndex = 8
        Me.btnBDashboard.Text = "Dashboard"
        Me.btnBDashboard.UseVisualStyleBackColor = False
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.btnBUsers)
        Me.Panel5.Location = New System.Drawing.Point(2, 210)
        Me.Panel5.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(144, 42)
        Me.Panel5.TabIndex = 1
        '
        'btnBUsers
        '
        Me.btnBUsers.BackColor = System.Drawing.Color.FromArgb(CType(CType(63, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(126, Byte), Integer))
        Me.btnBUsers.FlatAppearance.BorderSize = 0
        Me.btnBUsers.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBUsers.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBUsers.ForeColor = System.Drawing.Color.White
        Me.btnBUsers.Location = New System.Drawing.Point(-3, -1)
        Me.btnBUsers.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBUsers.Name = "btnBUsers"
        Me.btnBUsers.Size = New System.Drawing.Size(146, 41)
        Me.btnBUsers.TabIndex = 5
        Me.btnBUsers.Text = "Users"
        Me.btnBUsers.UseVisualStyleBackColor = False
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.Panel6)
        Me.Panel4.Controls.Add(Me.btnBBooks)
        Me.Panel4.Location = New System.Drawing.Point(2, 141)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(144, 42)
        Me.Panel4.TabIndex = 0
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(226, Byte), Integer), CType(CType(243, Byte), Integer))
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.ForeColor = System.Drawing.Color.Black
        Me.Panel6.Location = New System.Drawing.Point(10, 0)
        Me.Panel6.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(13, 42)
        Me.Panel6.TabIndex = 21
        '
        'btnBBooks
        '
        Me.btnBBooks.BackColor = System.Drawing.Color.FromArgb(CType(CType(63, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(126, Byte), Integer))
        Me.btnBBooks.FlatAppearance.BorderSize = 0
        Me.btnBBooks.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBBooks.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBBooks.ForeColor = System.Drawing.Color.White
        Me.btnBBooks.Location = New System.Drawing.Point(-3, -1)
        Me.btnBBooks.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBBooks.Name = "btnBBooks"
        Me.btnBBooks.Size = New System.Drawing.Size(146, 41)
        Me.btnBBooks.TabIndex = 4
        Me.btnBBooks.Text = "Books"
        Me.btnBBooks.UseVisualStyleBackColor = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Century Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(199, 77)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(54, 19)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Author"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Century Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(74, 77)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(71, 19)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Book Title"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(226, Byte), Integer), CType(CType(243, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.txtID)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.btnclear)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.txtBSearch)
        Me.Panel3.Controls.Add(Me.dgBooksList)
        Me.Panel3.Controls.Add(Me.Label11)
        Me.Panel3.Controls.Add(Me.btnBRefresh)
        Me.Panel3.Controls.Add(Me.btnBSave)
        Me.Panel3.Controls.Add(Me.btnBEdit)
        Me.Panel3.Controls.Add(Me.btnBDelete)
        Me.Panel3.Controls.Add(Me.btnBAdd)
        Me.Panel3.Controls.Add(Me.txtPrice)
        Me.Panel3.Controls.Add(Me.txtQuantity)
        Me.Panel3.Controls.Add(Me.cboCategories)
        Me.Panel3.Controls.Add(Me.txtAuthor)
        Me.Panel3.Controls.Add(Me.txtBTitle)
        Me.Panel3.Controls.Add(Me.Label10)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Location = New System.Drawing.Point(155, 9)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(665, 442)
        Me.Panel3.TabIndex = 7
        '
        'txtID
        '
        Me.txtID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtID.Font = New System.Drawing.Font("Century Gothic", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtID.Location = New System.Drawing.Point(10, 97)
        Me.txtID.Margin = New System.Windows.Forms.Padding(2)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(53, 25)
        Me.txtID.TabIndex = 25
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Century Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 77)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 19)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "Book ID"
        '
        'btnclear
        '
        Me.btnclear.BackColor = System.Drawing.Color.FromArgb(CType(CType(104, Byte), Integer), CType(CType(140, Byte), Integer), CType(CType(202, Byte), Integer))
        Me.btnclear.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnclear.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnclear.ForeColor = System.Drawing.Color.White
        Me.btnclear.Location = New System.Drawing.Point(496, 142)
        Me.btnclear.Margin = New System.Windows.Forms.Padding(2)
        Me.btnclear.Name = "btnclear"
        Me.btnclear.Size = New System.Drawing.Size(85, 23)
        Me.btnclear.TabIndex = 23
        Me.btnclear.Text = "Clear"
        Me.btnclear.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Century Schoolbook", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(209, 213)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 17)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Search"
        '
        'txtBSearch
        '
        Me.txtBSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBSearch.ForeColor = System.Drawing.Color.Black
        Me.txtBSearch.Location = New System.Drawing.Point(262, 211)
        Me.txtBSearch.Margin = New System.Windows.Forms.Padding(2)
        Me.txtBSearch.Multiline = True
        Me.txtBSearch.Name = "txtBSearch"
        Me.txtBSearch.Size = New System.Drawing.Size(180, 24)
        Me.txtBSearch.TabIndex = 21
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Century Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(330, 77)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(83, 19)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "Categories"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(63, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(126, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Panel8)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Panel7)
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Location = New System.Drawing.Point(4, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(154, 462)
        Me.Panel1.TabIndex = 6
        '
        'Panel8
        '
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel8.Controls.Add(Me.btnBBookshop)
        Me.Panel8.Location = New System.Drawing.Point(28, 79)
        Me.Panel8.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(99, 34)
        Me.Panel8.TabIndex = 23
        '
        'btnBBookshop
        '
        Me.btnBBookshop.BackColor = System.Drawing.Color.FromArgb(CType(CType(63, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(126, Byte), Integer))
        Me.btnBBookshop.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBBookshop.Font = New System.Drawing.Font("Century Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBBookshop.ForeColor = System.Drawing.Color.White
        Me.btnBBookshop.Location = New System.Drawing.Point(-2, -2)
        Me.btnBBookshop.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBBookshop.Name = "btnBBookshop"
        Me.btnBBookshop.Size = New System.Drawing.Size(99, 37)
        Me.btnBBookshop.TabIndex = 0
        Me.btnBBookshop.Text = "BookShop"
        Me.btnBBookshop.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(28, 9)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(99, 71)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.btnBLogout)
        Me.Panel2.Location = New System.Drawing.Point(13, 405)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(134, 32)
        Me.Panel2.TabIndex = 10
        '
        'btnBLogout
        '
        Me.btnBLogout.BackColor = System.Drawing.Color.FromArgb(CType(CType(63, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(126, Byte), Integer))
        Me.btnBLogout.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBLogout.Font = New System.Drawing.Font("Century Gothic", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBLogout.ForeColor = System.Drawing.Color.White
        Me.btnBLogout.Location = New System.Drawing.Point(15, -1)
        Me.btnBLogout.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBLogout.Name = "btnBLogout"
        Me.btnBLogout.Size = New System.Drawing.Size(102, 32)
        Me.btnBLogout.TabIndex = 9
        Me.btnBLogout.Text = "Log out"
        Me.btnBLogout.UseVisualStyleBackColor = False
        '
        'Books
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(63, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(126, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(831, 462)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Books"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Books"
        CType(Me.dgBooksList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel7.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel8.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgBooksList As DataGridView
    Friend WithEvents Label11 As Label
    Friend WithEvents txtPrice As TextBox
    Friend WithEvents cboCategories As ComboBox
    Friend WithEvents txtAuthor As TextBox
    Friend WithEvents txtBTitle As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents txtQuantity As TextBox
    Friend WithEvents btnBRefresh As Button
    Friend WithEvents btnBSave As Button
    Friend WithEvents btnBEdit As Button
    Friend WithEvents btnBDelete As Button
    Friend WithEvents btnBAdd As Button
    Friend WithEvents Label9 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Panel7 As Panel
    Friend WithEvents Panel5 As Panel
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label8 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnBBooks As Button
    Friend WithEvents btnBDashboard As Button
    Friend WithEvents btnBUsers As Button
    Friend WithEvents Panel6 As Panel
    Friend WithEvents txtBSearch As TextBox
    Friend WithEvents btnBLogout As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btnBBookshop As Button
    Friend WithEvents Panel8 As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents btnclear As Button
    Friend WithEvents txtID As TextBox
    Friend WithEvents Label2 As Label
End Class

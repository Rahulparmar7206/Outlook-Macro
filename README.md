# 📨 Outlook Macro — Save Multiple Attachments Easily

## 📘 Overview
This Outlook VBA Macro allows you to **save multiple email attachments** (or even single attachments) in one click.  
It includes a **UserForm** interface for selecting the save path and naming the output folder — making attachment saving fast, clean, and organized.

---

## ⚙️ Features
- 📁 Save **all attachments** from selected emails with a single click  
- 🧾 Choose **custom save location** via browse button  
- 🗂️ Automatically create a new folder for attachments  
- 💡 Simple and clean **VBA UserForm** interface  
- 🔒 Works entirely **offline** within Outlook — no add-ins required

---

## 🧩 UserForm Design (VBA)

| Control Type | Name | Caption / Label | Position / Size |
|---------------|-------|------------------|----------------|
| **Label1** | – | `Selected location: C:\Users\Admin\Desktop` | Top: 10, Left: 10, Width: 370, Height: 15 |
| **TextBox1** | `txtPath` | – | Top: 30, Left: 10, Width: 280, Height: 20 |
| **CommandButton1** | `btnBrowse` | `Browse for path` | Top: 30, Left: 300, Width: 80, Height: 22 |
| **Label2** | – | `Enter the folder name to save attachments:` | Top: 65, Left: 10, Width: 370, Height: 15 |
| **TextBox2** | `txtFolderName` | – | Top: 85, Left: 10, Width: 370, Height: 20 |
| **CommandButton2** | `btnOK` | `OK` | Top: 120, Left: 220, Width: 75, Height: 25 |
| **CommandButton3** | `btnCancel` | `Cancel` | Top: 120, Left: 305, Width: 75, Height: 25 |

---

## 💻 VBA Code Example

```vba
Sub SaveMailAttachments()
    ' Show UserForm
    UserForm1.Show
End Sub

Private Sub btnBrowse_Click()
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    If fldr.Show = -1 Then
        txtPath.Text = fldr.SelectedItems(1)
    End If
End Sub

Private Sub btnOK_Click()
    Dim savePath As String, folderName As String, itm As Object, atmt As Attachment
    savePath = txtPath.Text
    folderName = txtFolderName.Text
    
    If Right(savePath, 1) <> "\" Then savePath = savePath & "\"
    MkDir savePath & folderName
    
    For Each itm In Application.ActiveExplorer.Selection
        For Each atmt In itm.Attachments
            atmt.SaveAsFile savePath & folderName & "\" & atmt.FileName
        Next atmt
    Next itm
    
    MsgBox "Attachments saved successfully!", vbInformation
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
```

---

## 🪄 How to Use
1. Open **Outlook → Alt + F11** to open VBA Editor  
2. Insert a new **UserForm**  
3. Add the controls as per the above design  
4. Copy the VBA code into the UserForm code window  
5. Save and close the VBA editor  
6. Select one or more emails → Run macro → Choose folder → Done ✅

---

## 🧰 Requirements
- Microsoft Outlook (any desktop version)
- Macro security set to **Enable all macros**
- Windows system with access to file system

---

## 📜 License
This project is released under the **MIT License** — feel free to use and modify for your own workflow automation.

---

## 👨‍💻 Author
**Rahul Parmar**  
Finance Professional | Automation Enthusiast  
💼 GitHub: [@Rahulparmar7206](https://github.com/Rahulparmar7206)

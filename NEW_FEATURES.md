# App Update - New Features

## ✨ What's New

The app now has **TWO buttons** for generating Excel files, giving you complete flexibility:

### 1. 💾 **Save As New Excel** Button
- **Always visible** after data extraction
- Opens a file dialog to choose filename and location
- Use this for the **first save** or when you want a **different filename**

### 2. 🔄 **Regenerate Same File** Button
- **Appears after first save**
- Updates the same Excel file with new column parameters
- Asks for confirmation before overwriting
- Perfect for when you want to **update column values** (like FIRMA, KOSTSTELLE, etc.)

## 📋 How to Use

### Workflow Example:

1. **Load PDF & Extract Data**
   - Select PDF file
   - Click "Extract Data"
   - Preview window shows all entries

2. **First Time Save**
   - Configure your parameters (FIRMA, SOLL_HABEN, etc.)
   - Click **"💾 Save As New Excel"**
   - Choose filename (e.g., "November_Invoices.xlsx")
   - File is created ✅

3. **Update Columns & Regenerate**
   - Change some parameters (e.g., FIRMA from 9251 to 9252)
   - Click **"🔄 Regenerate Same File"**
   - Confirm overwrite
   - Same file is updated with new values ✅

4. **Save with Different Name (Optional)**
   - Change parameters again
   - Click **"💾 Save As New Excel"** again
   - Choose a new filename (e.g., "November_Invoices_v2.xlsx")
   - New file is created ✅

## 🎯 Key Benefits

✅ **Flexibility**: Choose new name or keep the same  
✅ **Safety**: Confirmation dialog before overwriting  
✅ **Convenience**: No need to type filename again for updates  
✅ **Visual Feedback**: Status label shows which file was saved  

## 📝 Example Scenario

**Scenario**: You realize you used the wrong KOSTSTELLE value

**Old way**: 
- Save file
- Close app
- Re-open app
- Load PDF again
- Extract again
- Change KOSTSTELLE
- Save with new name
- Delete old file

**New way**:
- Change KOSTSTELLE value in the form
- Click "🔄 Regenerate Same File"
- Done! ✅

## 🔍 Visual Indicators

After saving, you'll see:
```
💾 Saved: November_Invoices.xlsx
```

After regenerating:
```
🔄 Regenerated: November_Invoices.xlsx
```

This lets you know exactly what happened with your file!

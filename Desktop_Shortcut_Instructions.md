# Creating a Desktop Shortcut for Q3 Pacing Analysis

## Method 1: Create Desktop Shortcut (Recommended)

### Step 1: Right-click on Desktop
1. Right-click on an empty area of your desktop
2. Select **New** → **Shortcut**

### Step 2: Enter the Target Path
In the "Type the location of the item" field, enter:
```
C:\Users\harpe\OneDrive\Documents\VSCode\Pacing_Summary\Q3_Pacing_Analysis.bat
```

### Step 3: Name the Shortcut
1. Click **Next**
2. Enter a name for the shortcut: **Q3 Pacing Analysis**
3. Click **Finish**

### Step 4: Customize the Icon (Optional)
1. Right-click on the new desktop shortcut
2. Select **Properties**
3. Click **Change Icon...**
4. Choose an icon you like (or browse for a custom icon)
5. Click **OK** twice

## Method 2: Copy and Create Shortcut

### Alternative Method:
1. Navigate to: `C:\Users\harpe\OneDrive\Documents\VSCode\Pacing_Summary\`
2. Right-click on `Q3_Pacing_Analysis.bat`
3. Select **Send to** → **Desktop (create shortcut)**

## How to Use the Desktop Shortcut

1. **Prepare Input Files**: Place your Excel files in the `inputs` folder
2. **Double-click** the desktop shortcut
3. **Wait for Analysis**: The tool will automatically:
   - Find your input files
   - Load previous week data (if available)
   - Process all station data
   - Generate the output file
4. **Check Results**: Look in the `output` folder for your results

## What the Shortcut Does

When you double-click the desktop shortcut, it will:
- Open a command window with a professional title
- Show progress as it processes your data
- Display completion message
- Wait for you to press a key before closing

## Troubleshooting

**If the shortcut doesn't work:**
1. Verify the path in the shortcut properties
2. Make sure the `Q3_Pacing_Analysis.bat` file exists
3. Ensure your input files are in the `inputs` folder

**If you see errors:**
- Check that your Excel files have the correct date format (MM.DD.YY)
- Ensure you have at least 2 input files
- Verify the virtual environment is set up correctly

## File Locations

- **Script Location**: `C:\Users\harpe\OneDrive\Documents\VSCode\Pacing_Summary\`
- **Input Files**: `C:\Users\harpe\OneDrive\Documents\VSCode\Pacing_Summary\inputs\`
- **Output Files**: `C:\Users\harpe\OneDrive\Documents\VSCode\Pacing_Summary\output\`

Your desktop shortcut will make running the Q3 Pacing Analysis as easy as double-clicking an icon!

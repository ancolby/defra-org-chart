# DEFRA Senior Civil Service Organization Chart

An interactive web application that visualizes the UK Department for Environment, Food and Rural Affairs (DEFRA) Senior Civil Service organizational hierarchy. Features dynamic D3.js visualization with PowerPoint export capabilities.

## 🌟 Features

- **Interactive Organization Chart**: Hierarchical tree visualization with expand/collapse functionality
- **PowerPoint Export**: Generate professional .pptx files directly in the browser
- **Search & Filter**: Find personnel by name, position, or function
- **Grade Color Coding**: Visual distinction between SCS4-SCS1 levels
- **Responsive Design**: Works across desktop and mobile devices
- **Real-time Data Processing**: Handles complex CSV organizational data
- **Static Hosting Compatible**: No server required - runs entirely client-side

## 🎯 Live Demo

Open `index.html` in your browser to view the interactive organization chart.

## 📊 Screenshots

### Main Interface
- **Top**: Senior leadership overview (SCS4 Permanent Secretary + SCS3 Directors General)
- **Expandable Directorates**: Click any node to expand and view subordinates
- **Color-coded Hierarchy**: 
  - 🔴 **SCS4** (Dark Red): Permanent Secretary
  - 🔴 **SCS3** (Red): Directors General
  - 🔵 **SCS2** (Blue): Directors
  - 🟢 **SCS1** (Green): Deputy Directors

### PowerPoint Export
- **Multi-slide Output**: Overview + detailed directorate slides
- **Professional Formatting**: SmartArt-style organization charts
- **Current View Export**: Preserves collapsed/expanded states

## 🚀 Quick Start

### 1. Clone or Download
```bash
git clone <repository-url>
cd defra-org-chart
```

### 2. Open in Browser
Simply open `index.html` in any modern web browser. No build process or server required!

### 3. View Organization Chart
- The chart will automatically load DEFRA organizational data
- Use controls to expand/collapse sections
- Search for specific personnel or roles
- Click "📊 Export to PowerPoint" to download as .pptx

## 📁 Project Structure

```
├── index.html                          # Main application page
├── script.js                           # Chart visualization and interactions  
├── style.css                          # Styling and layout
├── data-processor.js                   # CSV parsing and hierarchy building
├── client-powerpoint-generator.js      # Browser PowerPoint generation
├── 2026-02-05T16-24-47Z-*.csv         # DEFRA organizational data
└── README.md                          # This file
```

## 🔧 Technical Details

### Dependencies
- **D3.js**: Data visualization and DOM manipulation
- **PptxGenJS**: Client-side PowerPoint generation
- **Modern Browser**: ES6+ JavaScript support required

### Browser Compatibility
- ✅ Chrome/Edge 80+
- ✅ Firefox 75+
- ✅ Safari 13+
- ❌ Internet Explorer (not supported)

### Data Format
The application expects CSV data with these key columns:
- `Post Unique Reference`: Unique identifier
- `Name`: Person's name (or "N/D" for undisclosed positions)
- `Grade (or equivalent)`: SCS4, SCS3, SCS2, SCS1
- `Job Title`: Position title
- `Job/Team Function`: Functional area
- `Unit`: Organizational unit/directorate
- `Reports to Senior Post`: Manager's unique reference (or "XX" for root)

### Architecture
1. **Data Processing**: CSV → Cleaned Objects → Hierarchical Tree
2. **Visualization**: D3.js renders SVG-based organization chart
3. **Interactivity**: Click handlers for expand/collapse and search
4. **Export**: Serialize current state → Generate PowerPoint in browser

## 📈 Usage

### Basic Navigation
- **Expand/Collapse**: Click any node with children to toggle visibility
- **Pan & Zoom**: Mouse drag to pan, scroll wheel to zoom
- **Search**: Type in search box to highlight matching personnel
- **Reset View**: Use "Expand All" / "Collapse All" buttons

### PowerPoint Export
1. **Set Desired View**: Expand or collapse sections as needed
2. **Click Export Button**: "📊 Export to PowerPoint" 
3. **Wait for Generation**: Takes 2-5 seconds for the library to load and generate
4. **Download**: File downloads automatically to your default folder

### Search Functionality
- Search across names, job titles, functions, and units
- Case-insensitive matching
- Results highlighted with golden outline
- Clear search by deleting text

## 🛠️ Customization

### Updating Data
Replace the CSV file with new organizational data. Ensure column headers match the expected format.

### Styling Changes
Edit `style.css` to modify:
- Colors and typography
- Layout dimensions
- Animation effects
- Responsive breakpoints

### Chart Configuration
In `script.js`, modify:
```javascript
// Chart dimensions
this.nodeWidth = 200;      // Node width in pixels
this.nodeHeight = 120;     // Node height in pixels
this.levelHeight = 180;    // Vertical spacing between levels
this.nodeSpacing = 250;    // Horizontal spacing between nodes
```

### PowerPoint Options
In `client-powerpoint-generator.js`:
```javascript
const generator = new ClientPowerPointGenerator({
    title: 'Custom Title',
    maxNodesPerSlide: 25,     // Nodes per slide before splitting
    colorScheme: 'defra',     // Color scheme
    slideWidth: 13.33,        // Slide width in inches
    slideHeight: 7.5          // Slide height in inches
});
```

## 🌐 Deployment

### Static Hosting (Recommended)
Works with any static hosting platform:

**GitHub Pages**:
```bash
git add .
git commit -m "Deploy organization chart"
git push origin main
# Enable Pages in repository settings
```

**Netlify**:
- Drag and drop folder to [netlify.com](https://netlify.com)
- Or connect GitHub repository for automatic deployments

**Azure Static Web Apps**:
- Create Static Web App resource
- Connect to GitHub repository
- Deploy automatically on commits

### Local Development
For local testing, simply open `index.html` in browser or use a local server:

```bash
# Python
python -m http.server 8000

# Node.js
npx serve .

# PHP
php -S localhost:8000
```

## 📊 Data Statistics

Current DEFRA organizational data includes:
- **214 personnel** across 4 SCS levels
- **7 directorates** (SCS3 organizational units)
- **4 organizational levels** (SCS4 → SCS3 → SCS2 → SCS1)
- **Complete hierarchy** with no orphaned records

## 🔍 Troubleshooting

### PowerPoint Export Issues
**"Library could not be loaded"**
- Check internet connection (PptxGenJS loads from CDN)
- Verify no ad blockers are preventing CDN access
- Try refreshing the page

**Download doesn't start**
- Check browser download settings
- Ensure pop-up blocker isn't active
- Verify browser supports file downloads

### Chart Loading Issues
**"Unable to Load Organizational Chart"**
- Ensure CSV file is in the same directory
- Check file permissions
- Verify CSV format matches expected structure

**Empty or malformed chart**
- Check browser console for JavaScript errors
- Verify CSV data has proper hierarchical relationships
- Ensure "XX" is used for root node in "Reports to" column

## 🤝 Contributing

1. **Fork** the repository
2. **Create** a feature branch: `git checkout -b feature-name`
3. **Commit** your changes: `git commit -m 'Add feature'`
4. **Push** to branch: `git push origin feature-name`
5. **Submit** a pull request

## 📜 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 📞 Support

For questions or issues:
1. **Check troubleshooting section** above
2. **Review browser console** for error messages
3. **Verify data format** matches requirements
4. **Test in different browser** if issues persist

## 🏢 About DEFRA

The Department for Environment, Food and Rural Affairs (DEFRA) is a UK government department responsible for environmental protection, food production and standards, agriculture, fisheries, and rural communities.

---

**Last Updated**: March 2026
**Data Source**: DEFRA Senior Civil Service Organogram
**Version**: 1.0.0
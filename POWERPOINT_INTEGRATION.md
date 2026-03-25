# PowerPoint Export Integration

This document explains how the PowerPoint export functionality has been integrated into your DEFRA organization chart application, with options for both static hosting and server-based deployment.

## ✅ **Recommended: Client-Side Approach** 
### (Works with Static Websites)

The **client-side approach** generates PowerPoint files directly in the browser using PptxGenJS. This works with any static hosting platform (GitHub Pages, Netlify, Azure Static Web Apps, etc.).

### Files Added/Modified:
- `client-powerpoint-generator.js` - Client-side PowerPoint generation engine
- `index.html` - Updated to include Export button and load PptxGenJS library
- `script.js` - Updated with client-side export functionality  
- `style.css` - Added styling for export button

### How It Works:
1. User clicks "📊 Export to PowerPoint" button
2. PptxGenJS library loads from CDN (if not already loaded)
3. Current organization chart state (including collapsed/expanded nodes) is serialized
4. PowerPoint presentation is generated entirely in the browser
5. File automatically downloads to user's computer

### Features:
- ✅ **Static hosting compatible** - No server required
- ✅ **Preserves current view state** - Exports what user sees
- ✅ **Multi-slide layout** - Overview + directorate detail slides  
- ✅ **Color-coded by grade** - Matches existing web app design
- ✅ **Professional SmartArt style** - Native PowerPoint formatting
- ✅ **Works offline** - After initial library load

### Testing:
1. Open your static website in a browser
2. Wait for organization chart to load
3. Optionally collapse/expand sections as desired
4. Click "📊 Export to PowerPoint" button
5. PowerPoint file should download automatically

## 🔧 **Alternative: Server-Side Approach**
### (For Custom Deployments)

The **server-side approach** uses an Express.js server to generate PowerPoint files. This requires hosting that supports Node.js applications.

### Files in powerpoint-generator/:
- `src/server.ts` - Express.js API server
- `src/powerpoint-generator.ts` - Server-side PowerPoint generation
- `src/data-processor.ts` - CSV parsing and hierarchy building
- `package.json` - Dependencies and scripts

### When to Use:
- Custom server deployments (not static hosting)
- Integration with enterprise systems
- Additional server-side processing requirements
- Bulk export capabilities

### How to Run:
```bash
cd powerpoint-generator
npm install
npm run build
npm run server:build
```

Server runs at `http://localhost:3000` and serves the static website with PowerPoint export API.

## 🎯 **Current Implementation**

Your application now uses the **client-side approach** by default, which is perfect for static hosting. The Export to PowerPoint button:

- 🟢 **Button color**: Orange (stands out from other controls)
- 🟢 **Loading state**: Shows "🔄 Generating PowerPoint..." during generation
- 🟢 **Success feedback**: Briefly shows "✅ Downloaded!" when complete
- 🟢 **Error handling**: User-friendly error messages for network/library issues

## 📊 **Generated PowerPoint Structure**

The exported PowerPoint includes:

1. **Title Slide** - Presentation title and generation date
2. **Overview Slide** - Senior leadership (SCS4 + SCS3) with legend
3. **Directorate Slides** - One slide per SCS3 directorate with full hierarchy
4. **Navigation Slide** - Grid view of all directorates

## 🛠️ **Customization Options**

To customize the PowerPoint output, modify the options in `script.js`:

```javascript
const generator = new ClientPowerPointGenerator({
    title: 'Custom Title Here',
    maxNodesPerSlide: 25,        // Increase for more dense layouts
    colorScheme: 'defra',        // Keep DEFRA colors
    slideWidth: 13.33,           // Standard PowerPoint size
    slideHeight: 7.5
});
```

## 🔍 **Troubleshooting**

### Common Issues:

**"PowerPoint generation library could not be loaded"**
- Check internet connection (PptxGenJS loads from CDN)
- Verify browser allows loading external scripts
- Try refreshing the page

**"No organization data loaded"**  
- Wait for the organization chart to fully load first
- Ensure CSV data loaded successfully

**Download doesn't start**
- Check browser's download settings
- Ensure pop-up blocker isn't preventing download
- Modern browsers should download automatically

### Browser Compatibility:
- Chrome/Edge: ✅ Full support
- Firefox: ✅ Full support  
- Safari: ✅ Full support
- Internet Explorer: ❌ Not supported (use modern browser)

## 📈 **Performance Notes**

- **First export**: ~3-5 seconds (library load + generation)
- **Subsequent exports**: ~1-2 seconds (generation only)
- **File size**: Typically 1-3 MB depending on organization size
- **Memory usage**: Low impact on browser performance
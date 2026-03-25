/**
 * Client-side PowerPoint Generator for Static Websites
 * Uses PptxGenJS in browser to generate PowerPoint files without requiring a server
 */

class ClientPowerPointGenerator {
    constructor(options = {}) {
        this.options = {
            title: 'DEFRA Senior Civil Service Organization Chart',
            maxNodesPerSlide: 20,
            colorScheme: 'defra',
            slideWidth: 13.33,
            slideHeight: 7.5,
            ...options
        };
        
        // Load PptxGenJS dynamically if not already loaded
        this.loadPptxGenJS();
    }

    async loadPptxGenJS() {
        // Check if PptxGenJS is already loaded
        if (typeof PptxGenJS !== 'undefined') {
            this.PptxGenJS = PptxGenJS;
            return;
        }

        try {
            // Load PptxGenJS from CDN
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
            script.onload = () => {
                this.PptxGenJS = window.PptxGenJS;
                console.log('✅ PptxGenJS loaded successfully');
            };
            script.onerror = () => {
                throw new Error('Failed to load PptxGenJS library');
            };
            document.head.appendChild(script);
            
            // Wait for script to load
            await new Promise((resolve, reject) => {
                script.onload = resolve;
                script.onerror = reject;
            });
            
        } catch (error) {
            console.error('❌ Failed to load PptxGenJS:', error);
            throw new Error('PowerPoint generation library could not be loaded. Please check your internet connection.');
        }
    }

    async generateFromHierarchy(rootNode, filename = null) {
        await this.loadPptxGenJS();
        
        if (!this.PptxGenJS) {
            throw new Error('PptxGenJS library not loaded');
        }

        console.log('🎯 Starting client-side PowerPoint generation...');
        
        // Initialize presentation
        const pptx = new this.PptxGenJS();
        
        // Configure presentation
        pptx.defineLayout({
            name: 'DEFRA_LAYOUT',
            width: this.options.slideWidth,
            height: this.options.slideHeight
        });
        pptx.layout = 'DEFRA_LAYOUT';
        
        // Set presentation metadata
        pptx.author = 'DEFRA Organization Chart Generator';
        pptx.company = 'Department for Environment, Food and Rural Affairs';
        pptx.subject = 'Senior Civil Service Organization Chart';
        pptx.title = this.options.title;

        // Create slides
        this.createTitleSlide(pptx);
        this.createOverviewSlide(pptx, rootNode);
        
        // Create directorate slides
        const directorates = this.findDirectorates(rootNode);
        directorates.forEach((directorate, index) => {
            console.log(`Creating slide for ${directorate.name} directorate (${index + 1}/${directorates.length})`);
            this.createDirectorateSlide(pptx, directorate, index + 3);
        });
        
        // Create navigation slide
        this.createNavigationSlide(pptx, directorates);
        
        // Generate filename if not provided
        const outputFilename = filename || `defra-org-chart-${new Date().toISOString().split('T')[0]}.pptx`;
        
        console.log('💾 Generating PowerPoint file for download...');
        
        // Generate and download the file
        await pptx.writeFile({ fileName: outputFilename });
        
        console.log(`✅ PowerPoint file generated: ${outputFilename}`);
        return outputFilename;
    }

    createTitleSlide(pptx) {
        const slide = pptx.addSlide();
        
        // Title
        slide.addText(this.options.title, {
            x: 1,
            y: 2,
            w: 11.33,
            h: 1.5,
            fontSize: 32,
            fontFace: 'Calibri',
            color: '1f497d',
            bold: true,
            align: 'center'
        });
        
        // Subtitle
        slide.addText('Senior Civil Service Hierarchy', {
            x: 1,
            y: 3.5,
            w: 11.33,
            h: 1,
            fontSize: 18,
            fontFace: 'Calibri',
            color: '666666',
            align: 'center'
        });
        
        // Date
        const currentDate = new Date().toLocaleDateString('en-GB', {
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        });
        
        slide.addText(`Generated on ${currentDate}`, {
            x: 1,
            y: 6,
            w: 11.33,
            h: 0.5,
            fontSize: 12,
            fontFace: 'Calibri',
            color: '999999',
            align: 'center'
        });
    }

    createOverviewSlide(pptx, rootNode) {
        const slide = pptx.addSlide();
        
        slide.addText('Senior Leadership Overview', {
            x: 0.5,
            y: 0.3,
            w: 12.33,
            h: 0.8,
            fontSize: 24,
            fontFace: 'Calibri',
            color: '1f497d',
            bold: true,
            align: 'center'
        });
        
        // Create simplified org chart for top levels
        this.createOrgChart(slide, rootNode, {
            startX: 1,
            startY: 1.5,
            width: 11.33,
            height: 5,
            maxLevels: 2,
            nodeWidth: 2.2,
            nodeHeight: 1,
            levelSpacing: 1.4,
            siblingSpacing: 0.3
        });
        
        // Add legend
        this.addGradeLegend(slide, 0.5, 6.5);
    }

    createDirectorateSlide(pptx, directorate, slideNumber) {
        const slide = pptx.addSlide();
        
        const title = `${directorate.name} - ${directorate.unit}`;
        slide.addText(title, {
            x: 0.5,
            y: 0.3,
            w: 12.33,
            h: 0.8,
            fontSize: 20,
            fontFace: 'Calibri',
            color: '1f497d',
            bold: true,
            align: 'center'
        });
        
        // Count nodes in this subtree
        const totalNodes = this.countNodes(directorate);
        
        if (totalNodes <= this.options.maxNodesPerSlide) {
            this.createOrgChart(slide, directorate, {
                startX: 0.5,
                startY: 1.5,
                width: 12.33,
                height: 5.5,
                maxLevels: 4,
                nodeWidth: 1.8,
                nodeHeight: 0.8,
                levelSpacing: 1.2,
                siblingSpacing: 0.2
            });
        } else {
            // Large directorate - show simplified view
            this.createLargeDirectorateView(slide, directorate);
        }
        
        // Add slide number
        slide.addText(`Slide ${slideNumber}`, {
            x: 11.5,
            y: 6.8,
            w: 1.5,
            h: 0.3,
            fontSize: 10,
            fontFace: 'Calibri',
            color: '666666',
            align: 'right'
        });
    }

    createOrgChart(slide, rootNode, layout) {
        const positions = this.calculatePositions(rootNode, layout);
        
        // Draw connections
        this.drawConnections(slide, positions);
        
        // Draw nodes
        positions.forEach(({ node, x, y }) => {
            const color = this.getGradeColor(node.grade);
            
            // Add rectangle
            slide.addShape('rect', {
                x: x,
                y: y,
                w: layout.nodeWidth,
                h: layout.nodeHeight,
                fill: { color: color },
                line: { color: 'ffffff', width: 2 },
                rectRadius: 0.1
            });
            
            // Add name
            slide.addText(node.name, {
                x: x + 0.1,
                y: y + 0.05,
                w: layout.nodeWidth - 0.2,
                h: layout.nodeHeight * 0.4,
                fontSize: Math.min(10, Math.max(8, layout.nodeWidth * 4)),
                fontFace: 'Calibri',
                color: 'ffffff',
                bold: true,
                align: 'center',
                valign: 'top'
            });
            
            // Add grade
            slide.addText(node.grade, {
                x: x + 0.1,
                y: y + layout.nodeHeight * 0.35,
                w: layout.nodeWidth - 0.2,
                h: layout.nodeHeight * 0.25,
                fontSize: Math.min(8, Math.max(6, layout.nodeWidth * 3)),
                fontFace: 'Calibri',
                color: 'ffffff',
                align: 'center',
                valign: 'middle'
            });
            
            // Add job title if space allows
            if (layout.nodeHeight > 0.6 && node.jobTitle) {
                slide.addText(node.jobTitle, {
                    x: x + 0.1,
                    y: y + layout.nodeHeight * 0.6,
                    w: layout.nodeWidth - 0.2,
                    h: layout.nodeHeight * 0.35,
                    fontSize: Math.min(7, Math.max(5, layout.nodeWidth * 2.5)),
                    fontFace: 'Calibri',
                    color: 'ffffff',
                    italic: true,
                    align: 'center',
                    valign: 'middle'
                });
            }
        });
    }

    calculatePositions(rootNode, layout, level = 0, positions = [], parentX = null) {
        if (level >= layout.maxLevels) return positions;
        
        const y = layout.startY + level * layout.levelSpacing;
        
        if (level === 0) {
            // Root node
            const x = layout.startX + layout.width / 2 - layout.nodeWidth / 2;
            positions.push({ node: rootNode, x, y, level });
            
            // Process children
            this.layoutChildren(rootNode, x + layout.nodeWidth / 2, level + 1, layout, positions);
        }
        
        return positions;
    }

    layoutChildren(parent, parentCenterX, level, layout, positions) {
        if (level >= layout.maxLevels || !parent.children || parent.children.length === 0) {
            return;
        }
        
        // Filter children based on collapsed state
        const visibleChildren = parent._collapsed ? [] : parent.children;
        
        if (visibleChildren.length === 0) return;
        
        const y = layout.startY + level * layout.levelSpacing;
        const totalWidth = visibleChildren.length * layout.nodeWidth + (visibleChildren.length - 1) * layout.siblingSpacing;
        const startX = parentCenterX - totalWidth / 2;
        
        visibleChildren.forEach((child, index) => {
            const x = startX + index * (layout.nodeWidth + layout.siblingSpacing);
            positions.push({ node: child, x, y, level });
            
            // Recursively layout grandchildren
            this.layoutChildren(child, x + layout.nodeWidth / 2, level + 1, layout, positions);
        });
    }

    drawConnections(slide, positions) {
        const positionMap = new Map();
        positions.forEach(pos => {
            positionMap.set(pos.node.id, pos);
        });
        
        positions.forEach(({ node, x, y, level }, index) => {
            if (node.children && !node._collapsed && node.children.length > 0) {
                const parentCenterX = x + 0.9;
                const parentBottomY = y + 0.8;
                
                node.children.forEach(child => {
                    const childPos = positionMap.get(child.id);
                    if (childPos) {
                        const childCenterX = childPos.x + 0.9;
                        const childTopY = childPos.y;
                        
                        // Draw vertical line from parent
                        slide.addShape('line', {
                            x: parentCenterX,
                            y: parentBottomY,
                            w: 0,
                            h: (childTopY - parentBottomY) / 2,
                            line: { color: '666666', width: 1 }
                        });
                        
                        // Draw horizontal line
                        slide.addShape('line', {
                            x: Math.min(parentCenterX, childCenterX),
                            y: parentBottomY + (childTopY - parentBottomY) / 2,
                            w: Math.abs(childCenterX - parentCenterX),
                            h: 0,
                            line: { color: '666666', width: 1 }
                        });
                        
                        // Draw vertical line to child
                        slide.addShape('line', {
                            x: childCenterX,
                            y: parentBottomY + (childTopY - parentBottomY) / 2,
                            w: 0,
                            h: (childTopY - parentBottomY) / 2,
                            line: { color: '666666', width: 1 }
                        });
                    }
                });
            }
        });
    }

    createLargeDirectorateView(slide, directorate) {
        slide.addText('Large Directorate - Showing Direct Reports Only', {
            x: 0.5,
            y: 1.2,
            w: 12.33,
            h: 0.4,
            fontSize: 12,
            fontFace: 'Calibri',
            color: '666666',
            italic: true,
            align: 'center'
        });
        
        this.createOrgChart(slide, directorate, {
            startX: 1,
            startY: 1.8,
            width: 11.33,
            height: 4.5,
            maxLevels: 2,
            nodeWidth: 2,
            nodeHeight: 0.8,
            levelSpacing: 1.5,
            siblingSpacing: 0.3
        });
        
        const totalNodes = this.countNodes(directorate);
        slide.addText(
            `Note: This directorate contains ${totalNodes} total positions. Only direct reports are shown for clarity.`,
            {
                x: 1,
                y: 6.5,
                w: 11.33,
                h: 0.5,
                fontSize: 10,
                fontFace: 'Calibri',
                color: '666666',
                align: 'center'
            }
        );
    }

    createNavigationSlide(pptx, directorates) {
        const slide = pptx.addSlide();
        
        slide.addText('Navigation - Directorate Overview', {
            x: 1,
            y: 0.5,
            w: 11.33,
            h: 1,
            fontSize: 24,
            fontFace: 'Calibri',
            color: '1f497d',
            bold: true,
            align: 'center'
        });
        
        const cols = 3;
        const rows = Math.ceil(directorates.length / cols);
        const boxWidth = 3.5;
        const boxHeight = 1.2;
        const marginX = 0.4;
        const marginY = 0.3;
        
        directorates.forEach((directorate, index) => {
            const row = Math.floor(index / cols);
            const col = index % cols;
            
            const x = 1 + col * (boxWidth + marginX);
            const y = 2 + row * (boxHeight + marginY);
            
            // Add colored box
            slide.addShape('rect', {
                x,
                y,
                w: boxWidth,
                h: boxHeight,
                fill: { color: this.getGradeColor(directorate.grade) },
                line: { color: 'ffffff', width: 2 },
                rectRadius: 0.1
            });
            
            // Add name
            slide.addText(directorate.name, {
                x: x + 0.1,
                y: y + 0.1,
                w: boxWidth - 0.2,
                h: boxHeight * 0.6,
                fontSize: 12,
                fontFace: 'Calibri',
                color: 'ffffff',
                bold: true,
                align: 'center',
                valign: 'middle'
            });
            
            // Add unit
            slide.addText(directorate.unit, {
                x: x + 0.1,
                y: y + boxHeight * 0.6,
                w: boxWidth - 0.2,
                h: boxHeight * 0.4,
                fontSize: 9,
                fontFace: 'Calibri',
                color: 'ffffff',
                align: 'center',
                valign: 'middle'
            });
        });
    }

    addGradeLegend(slide, x, y) {
        const grades = [
            { grade: 'SCS4', label: 'Permanent Secretary', color: this.getGradeColor('SCS4') },
            { grade: 'SCS3', label: 'Directors General', color: this.getGradeColor('SCS3') },
            { grade: 'SCS2', label: 'Directors', color: this.getGradeColor('SCS2') },
            { grade: 'SCS1', label: 'Deputy Directors', color: this.getGradeColor('SCS1') }
        ];
        
        slide.addText('Grade Legend', {
            x: x,
            y: y,
            w: 2,
            h: 0.3,
            fontSize: 12,
            fontFace: 'Calibri',
            color: '1f497d',
            bold: true
        });
        
        grades.forEach((gradeInfo, index) => {
            const legendY = y + 0.4 + index * 0.4;
            
            // Color box
            slide.addShape('rect', {
                x: x,
                y: legendY,
                w: 0.3,
                h: 0.3,
                fill: { color: gradeInfo.color },
                line: { color: 'cccccc', width: 1 }
            });
            
            // Label
            slide.addText(`${gradeInfo.grade} - ${gradeInfo.label}`, {
                x: x + 0.4,
                y: legendY,
                w: 2.5,
                h: 0.3,
                fontSize: 9,
                fontFace: 'Calibri',
                color: '333333',
                valign: 'middle'
            });
        });
    }

    findDirectorates(rootNode) {
        const directorates = [];
        
        function findSCS3Nodes(node) {
            if (node.grade === 'SCS3') {
                directorates.push(node);
            }
            if (node.children && !node._collapsed) {
                node.children.forEach(child => findSCS3Nodes(child));
            }
        }
        
        findSCS3Nodes(rootNode);
        return directorates;
    }

    countNodes(node) {
        let count = 1;
        if (node.children && !node._collapsed) {
            node.children.forEach(child => {
                count += this.countNodes(child);
            });
        }
        return count;
    }

    getGradeColor(grade) {
        const colorMap = {
            'SCS4': '8b0000', // Dark Red
            'SCS3': 'e74c3c', // Bright Red
            'SCS2': '3498db', // Blue
            'SCS1': '27ae60'  // Green
        };
        return colorMap[grade] || '6c757d'; // Default gray
    }
}
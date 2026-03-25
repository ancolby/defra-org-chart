class OrgChart {
    constructor() {
        this.svg = d3.select('#org-chart');
        this.g = this.svg.append('g');
        this.tooltip = d3.select('#tooltip');
        
        // Chart dimensions and settings
        this.nodeWidth = 200;
        this.nodeHeight = 120;
        this.levelHeight = 180;
        this.nodeSpacing = 250;
        
        // Data
        this.root = null;
        this.nodeMap = new Map();
        
        // Setup zoom behavior
        this.setupZoom();
        this.setupEventListeners();
        
        // Load data
        this.loadData();
    }
    
    setupZoom() {
        const zoom = d3.zoom()
            .scaleExtent([0.1, 3])
            .on('zoom', (event) => {
                this.g.attr('transform', event.transform);
            });
        
        this.svg.call(zoom);
        
        // Set initial zoom to center
        const initialTransform = d3.zoomIdentity.translate(50, 50).scale(0.8);
        this.svg.call(zoom.transform, initialTransform);
    }
    
    setupEventListeners() {
        // Expand/Collapse buttons
        d3.select('#expand-all').on('click', () => this.expandAll());
        d3.select('#collapse-all').on('click', () => this.collapseAll());
        
        // Export to PowerPoint button
        d3.select('#export-powerpoint').on('click', () => this.exportToPowerPoint());
        
        // Search functionality
        d3.select('#search').on('input', (event) => {
            this.search(event.target.value);
        });
    }
    
    async loadData() {
        const loadingElement = document.getElementById('loading');
        
        try {
            // Show loading indicator
            if (loadingElement) {
                loadingElement.style.display = 'block';
            }
            
            // Load CSV file - using the exact filename from your directory
            const csvText = await d3.text('2026-02-05T16-24-47Z-252026-02-05-organogram-senior.csv');
            const rawData = DataProcessor.parseCSV(csvText);
            const cleanedData = DataProcessor.cleanData(rawData);
            
            console.log('Loaded data:', cleanedData.length, 'records');
            
            // Build hierarchy
            const { root, orphans, nodeMap } = DataProcessor.buildHierarchy(cleanedData);
            this.root = root;
            this.nodeMap = nodeMap;
            
            if (!root) {
                throw new Error('No root node found. Expected someone reporting to "XX"');
            }
            
            console.log('Root node:', root.name, root.grade);
            console.log('Orphaned nodes:', orphans.length);
            
            // Validate hierarchy
            const issues = DataProcessor.validateHierarchy(root, nodeMap);
            if (issues.length > 0) {
                console.warn('Hierarchy issues:', issues);
            }
            
            // Get stats
            const stats = DataProcessor.getNodeStats(root);
            console.log('Total nodes:', stats.totalNodes);
            console.log('Grade distribution:', stats.gradeCount);
            
            // Initialize node states
            this.initializeNodeStates(root);
            
            // Hide loading indicator
            if (loadingElement) {
                loadingElement.style.display = 'none';
            }
            
            // Render chart
            this.render();
            
        } catch (error) {
            console.error('Error loading data:', error);
            
            // Hide loading indicator
            if (loadingElement) {
                loadingElement.style.display = 'none';
            }
            
            // Show user-friendly error message
            this.showError(error);
        }
    }
    
    initializeNodeStates(node) {
        node._collapsed = false;
        node.children.forEach(child => this.initializeNodeStates(child));
    }
    
    render() {
        if (!this.root) return;
        
        // Calculate layout
        this.calculateLayout(this.root, 0, 0);
        
        // Get visible nodes and links
        const { nodes, links } = this.getVisibleNodesAndLinks(this.root);
        
        // Render links first (so they appear behind nodes)
        this.renderLinks(links);
        
        // Render nodes
        this.renderNodes(nodes);
    }
    
    calculateLayout(node, level, parentX = 0) {
        // Get visible children
        const visibleChildren = node._collapsed ? [] : node.children;
        
        // Calculate positions for children
        const childrenWidth = visibleChildren.length * this.nodeSpacing;
        const startX = parentX - childrenWidth / 2 + this.nodeSpacing / 2;
        
        // Position current node
        node.x = parentX;
        node.y = level * this.levelHeight;
        
        // Position children
        visibleChildren.forEach((child, index) => {
            const childX = startX + index * this.nodeSpacing;
            this.calculateLayout(child, level + 1, childX);
        });
    }
    
    getVisibleNodesAndLinks(root) {
        const nodes = [];
        const links = [];
        
        function traverse(node) {
            nodes.push(node);
            
            const visibleChildren = node._collapsed ? [] : node.children;
            
            visibleChildren.forEach(child => {
                links.push({ source: node, target: child });
                traverse(child);
            });
        }
        
        traverse(root);
        return { nodes, links };
    }
    
    renderLinks(links) {
        const linkSelection = this.g.selectAll('.link')
            .data(links, d => `${d.source.id}-${d.target.id}`);
        
        linkSelection.exit().remove();
        
        linkSelection.enter()
            .append('path')
            .attr('class', 'link')
            .merge(linkSelection)
            .attr('d', d => {
                const sourceX = d.source.x;
                const sourceY = d.source.y + this.nodeHeight;
                const targetX = d.target.x;
                const targetY = d.target.y;
                
                const midY = sourceY + (targetY - sourceY) / 2;
                
                return `M${sourceX},${sourceY} 
                        C${sourceX},${midY} 
                         ${targetX},${midY} 
                         ${targetX},${targetY}`;
            });
    }
    
    renderNodes(nodes) {
        const nodeSelection = this.g.selectAll('.node')
            .data(nodes, d => d.id);
        
        nodeSelection.exit().remove();
        
        const nodeEnter = nodeSelection.enter()
            .append('g')
            .attr('class', 'node')
            .style('cursor', 'pointer');
        
        const nodeUpdate = nodeEnter.merge(nodeSelection);
        
        nodeUpdate
            .transition()
            .duration(500)
            .attr('transform', d => `translate(${d.x - this.nodeWidth/2}, ${d.y})`);
        
        // Add/update rectangles
        const rectSelection = nodeUpdate.selectAll('.node-rect')
            .data(d => [d]);
        
        rectSelection.enter()
            .append('rect')
            .attr('class', d => `node-rect ${d.grade.toLowerCase()}`)
            .attr('width', this.nodeWidth)
            .attr('height', this.nodeHeight)
            .attr('rx', 8)
            .merge(rectSelection);
        
        // Add/update text elements
        this.addText(nodeUpdate, 'name', 15, d => d.name);
        this.addText(nodeUpdate, 'grade', 32, d => d.grade);
        this.addText(nodeUpdate, 'title', 50, d => d.jobTitle);
        this.addText(nodeUpdate, 'function', 68, d => d.function);
        this.addText(nodeUpdate, 'unit', 85, d => this.truncateText(d.unit, 25));
        
        // Add collapse indicators for nodes with children
        this.addCollapseIndicator(nodeUpdate);
        
        // Add event listeners
        nodeUpdate
            .on('click', (event, d) => {
                if (d.children.length > 0) {
                    d._collapsed = !d._collapsed;
                    this.render();
                }
            })
            .on('mouseover', (event, d) => {
                this.showTooltip(event, d);
            })
            .on('mouseout', () => {
                this.hideTooltip();
            });
    }
    
    addText(selection, className, y, textFn) {
        const textSelection = selection.selectAll(`.node-text.${className}`)
            .data(d => [d]);
        
        textSelection.enter()
            .append('text')
            .attr('class', `node-text ${className}`)
            .attr('x', this.nodeWidth / 2)
            .attr('y', y)
            .attr('dy', '0.35em')
            .merge(textSelection)
            .text(textFn);
    }
    
    addCollapseIndicator(selection) {
        // Remove existing indicators
        selection.selectAll('.collapse-indicator, .collapse-text').remove();
        
        // Add new indicators for nodes with children
        const nodesWithChildren = selection.filter(d => d.children.length > 0);
        
        nodesWithChildren.append('circle')
            .attr('class', 'collapse-indicator')
            .attr('cx', this.nodeWidth - 15)
            .attr('cy', 15)
            .attr('r', 8);
        
        nodesWithChildren.append('text')
            .attr('class', 'collapse-text')
            .attr('x', this.nodeWidth - 15)
            .attr('y', 15)
            .attr('dy', '0.35em')
            .text(d => d._collapsed ? '+' : '-');
    }
    
    truncateText(text, maxLength) {
        if (!text) return '';
        return text.length > maxLength ? text.substring(0, maxLength - 3) + '...' : text;
    }
    
    showTooltip(event, d) {
        const html = `
            <strong>${d.name}</strong><br>
            Grade: ${d.grade}<br>
            Position: ${d.jobTitle}<br>
            Function: ${d.function || 'Not specified'}<br>
            Unit: ${d.unit}<br>
            Reports to: ${d.reportsTo}<br>
            Direct Reports: ${d.children.length}
        `;
        
        this.tooltip
            .html(html)
            .style('left', (event.pageX + 10) + 'px')
            .style('top', (event.pageY - 10) + 'px')
            .style('opacity', 1);
    }
    
    hideTooltip() {
        this.tooltip.style('opacity', 0);
    }
    
    expandAll() {
        function expand(node) {
            node._collapsed = false;
            node.children.forEach(expand);
        }
        expand(this.root);
        this.render();
    }
    
    collapseAll() {
        function collapse(node) {
            if (node.children.length > 0) {
                node._collapsed = true;
            }
            node.children.forEach(collapse);
        }
        collapse(this.root);
        this.render();
    }
    
    search(term) {
        // Reset previous highlights
        this.g.selectAll('.node-rect').classed('search-highlight', false);
        
        if (!term.trim()) return;
        
        const searchTerm = term.toLowerCase();
        
        // Find matching nodes
        const matches = [];
        function findMatches(node) {
            const searchableText = [
                node.name,
                node.jobTitle,
                node.function,
                node.unit,
                node.grade
            ].join(' ').toLowerCase();
            
            if (searchableText.includes(searchTerm)) {
                matches.push(node);
            }
            
            node.children.forEach(findMatches);
        }
        
        findMatches(this.root);
        
        // Highlight matches
        matches.forEach(node => {
            this.g.selectAll('.node-rect')
                .filter(d => d.id === node.id)
                .classed('search-highlight', true);
        });
        
        console.log(`Found ${matches.length} matches for "${term}"`);
    }
    
    showError(error) {
        const chartContainer = document.getElementById('chart-container');
        
        const errorDiv = document.createElement('div');
        errorDiv.className = 'error-message';
        errorDiv.innerHTML = `
            <h3>Unable to Load Organizational Chart</h3>
            <p>Error: ${error.message}</p>
            <p>Please ensure all files are uploaded to the same folder and try refreshing the page.</p>
        `;
        
        chartContainer.appendChild(errorDiv);
    }
    
    async exportToPowerPoint() {
        if (!this.root) {
            alert('No organization data loaded. Please wait for the chart to load first.');
            return;
        }
        
        try {
            // Show loading state
            const exportButton = document.getElementById('export-powerpoint');
            const originalText = exportButton.textContent;
            exportButton.textContent = '🔄 Generating PowerPoint...';
            exportButton.disabled = true;
            
            console.log('📤 Starting client-side PowerPoint generation...');
            
            // Initialize client-side PowerPoint generator
            const generator = new ClientPowerPointGenerator({
                title: 'DEFRA Senior Civil Service Organization Chart',
                maxNodesPerSlide: 20,
                colorScheme: 'defra'
            });
            
            // Serialize current hierarchy state (preserving collapsed/expanded state)
            const hierarchyData = this.serializeHierarchy(this.root);
            
            // Generate PowerPoint file (will automatically download)
            const filename = await generator.generateFromHierarchy(hierarchyData);
            
            console.log('✅ PowerPoint export completed successfully:', filename);
            
            // Show success message briefly
            exportButton.textContent = '✅ Downloaded!';
            setTimeout(() => {
                exportButton.textContent = originalText;
            }, 2000);
            
        } catch (error) {
            console.error('❌ PowerPoint export failed:', error);
            
            // Show user-friendly error message
            let message = 'Failed to export PowerPoint. ';
            if (error.message.includes('PptxGenJS')) {
                message += 'Could not load PowerPoint generation library. Please check your internet connection and try again.';
            } else {
                message += error.message;
            }
            
            alert(message);
            
        } finally {
            // Restore button state
            const exportButton = document.getElementById('export-powerpoint');
            if (exportButton.textContent.includes('🔄')) {
                exportButton.textContent = originalText;
            }
            exportButton.disabled = false;
        }
    }
    
    serializeHierarchy(node) {
        if (!node) return null;
        
        // Create a serialized version of the node with current state
        const serialized = {
            id: node.id,
            name: node.name,
            grade: node.grade,
            jobTitle: node.jobTitle,
            function: node.function,
            unit: node.unit,
            reportsTo: node.reportsTo,
            level: node.level || 0,
            _collapsed: node._collapsed || false, // Preserve current collapsed state
            originalData: node.originalData,
            children: []
        };
        
        // Recursively serialize children
        if (node.children && node.children.length > 0) {
            serialized.children = node.children.map(child => this.serializeHierarchy(child));
        }
        
        return serialized;
    }
}

// Initialize chart when page loads
document.addEventListener('DOMContentLoaded', () => {
    new OrgChart();
});
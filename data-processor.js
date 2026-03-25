class DataProcessor {
    static parseCSV(csvText) {
        const lines = csvText.split('\n');
        const headers = this.parseCSVRow(lines[0]);
        const data = [];
        
        for (let i = 1; i < lines.length; i++) {
            if (lines[i].trim()) {
                const values = this.parseCSVRow(lines[i]);
                if (values.length === headers.length) {
                    const row = {};
                    headers.forEach((header, index) => {
                        row[header] = values[index];
                    });
                    data.push(row);
                }
            }
        }
        
        return data;
    }
    
    static parseCSVRow(row) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < row.length; i++) {
            const char = row[i];
            const nextChar = row[i + 1];
            
            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    current += '"';
                    i++; // Skip next quote
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        result.push(current.trim());
        return result;
    }
    
    static cleanData(data) {
        return data.map(row => ({
            id: row['Post Unique Reference'],
            name: row['Name'] === 'N/D' ? 'Not Disclosed' : row['Name'],
            grade: row['Grade (or equivalent)'],
            jobTitle: row['Job Title'],
            function: row['Job/Team Function'],
            unit: row['Unit'],
            reportsTo: row['Reports to Senior Post'],
            originalData: row
        }));
    }
    
    static buildHierarchy(cleanedData) {
        // Create lookup map
        const nodeMap = new Map();
        cleanedData.forEach(person => {
            nodeMap.set(person.id, {
                ...person,
                children: []
            });
        });
        
        // Find root and build hierarchy
        let root = null;
        const orphans = [];
        
        cleanedData.forEach(person => {
            const node = nodeMap.get(person.id);
            
            if (person.reportsTo === 'XX' || !person.reportsTo) {
                // This is the root
                root = node;
            } else {
                const parent = nodeMap.get(person.reportsTo);
                if (parent) {
                    parent.children.push(node);
                } else {
                    // Orphaned node - parent not found
                    orphans.push(node);
                }
            }
        });
        
        // Log any orphaned nodes for debugging
        if (orphans.length > 0) {
            console.warn('Orphaned nodes found:', orphans.map(n => `${n.name} (${n.id}) reports to ${n.reportsTo}`));
        }
        
        return { root, orphans, nodeMap };
    }
    
    static validateHierarchy(root, nodeMap) {
        const issues = [];
        const visited = new Set();
        
        function checkNode(node, path = []) {
            if (visited.has(node.id)) {
                issues.push({
                    type: 'circular',
                    message: `Circular reference detected: ${path.join(' -> ')} -> ${node.id}`
                });
                return;
            }
            
            visited.add(node.id);
            const newPath = [...path, node.id];
            
            node.children.forEach(child => {
                checkNode(child, newPath);
            });
            
            visited.delete(node.id);
        }
        
        if (root) {
            checkNode(root);
        }
        
        return issues;
    }
    
    static getNodeStats(root) {
        let totalNodes = 0;
        const gradeCount = {};
        
        function countNodes(node) {
            totalNodes++;
            gradeCount[node.grade] = (gradeCount[node.grade] || 0) + 1;
            
            node.children.forEach(child => countNodes(child));
        }
        
        if (root) {
            countNodes(root);
        }
        
        return { totalNodes, gradeCount };
    }
}
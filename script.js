document.addEventListener('DOMContentLoaded', () => {
    // Metric labels for display
    const metricLabels = {
        'mobileSales': 'Mobile Sales',
        'fiberSales': 'Fiber Sales',
        'promoter': 'Promoter',
        'videoPackage': 'Video Package'
    };
    const fileInput = document.getElementById('fileInput');
    const managerFilter = document.getElementById('managerFilter');
    const metricFilter = document.getElementById('metricFilter');
    const managerCards = document.getElementById('managerCards');
    const topManagerInfo = document.getElementById('topManagerInfo');
    
    let employeeData = [];
    let managers = new Set();
    
    // Handle file upload
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        try {
            const data = await readExcelFile(file);
            processEmployeeData(data);
            updateDashboard();
        } catch (error) {
            console.error('Error processing file:', error);
            alert('Error processing file. Please make sure it is a valid Excel file.');
        }
    });
    
    // Filter event listeners
    managerFilter.addEventListener('change', updateDashboard);
    metricFilter.addEventListener('change', updateDashboard);
    
    // Read Excel file
    function readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    
                    // Convert to array of objects
                    const headers = jsonData[0];
                    const result = jsonData.slice(1).map(row => {
                        const obj = {};
                        headers.forEach((header, index) => {
                            obj[header] = row[index] || 0;
                        });
                        return obj;
                    });
                    
                    resolve(result);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = (error) => reject(error);
            reader.readAsArrayBuffer(file);
        });
    }
    
    // Process employee data
    function processEmployeeData(data) {
        employeeData = data;
        managers = new Set(employeeData.map(emp => emp.Manager));
        
        // Update manager filter
        managerFilter.innerHTML = '<option value="">All Managers</option>';
        managers.forEach(manager => {
            if (manager) {
                const option = document.createElement('option');
                option.value = manager;
                option.textContent = manager;
                managerFilter.appendChild(option);
            }
        });
    }
    
    // Calculate total score for an employee across all metrics
    function calculateTotalScore(employee) {
        const metrics = ['mobileSales', 'fiberSales', 'promoter', 'videoPackage'];
        return metrics.reduce((sum, metric) => sum + (parseFloat(employee[metric]) || 0), 0);
    }

    // Update dashboard with filtered data
    function updateDashboard() {
        const selectedManager = managerFilter.value;
        const selectedMetric = metricFilter.value;
        
        // Filter data
        let filteredData = [...employeeData];
        if (selectedManager) {
            filteredData = filteredData.filter(emp => emp.Manager === selectedManager);
        }
        
        // Group by manager
        const managerGroups = filteredData.reduce((groups, emp) => {
            if (!emp.Manager) return groups;
            
            if (!groups[emp.Manager]) {
                groups[emp.Manager] = [];
            }
            
            groups[emp.Manager].push(emp);
            return groups;
        }, {});
        
        // Sort employees within each manager group
        Object.keys(managerGroups).forEach(manager => {
            managerGroups[manager].sort((a, b) => {
                if (selectedMetric === 'all') {
                    return calculateTotalScore(b) - calculateTotalScore(a);
                } else {
                    return (parseFloat(b[selectedMetric]) || 0) - (parseFloat(a[selectedMetric]) || 0);
                }
            });
        });
        
        // Sort managers by their team's performance
        const sortedManagers = Object.keys(managerGroups).sort((a, b) => {
            let topScoreA, topScoreB;
            
            if (selectedMetric === 'all') {
                // For 'All Metrics', calculate average of all metrics for top 3 employees
                topScoreA = managerGroups[a].slice(0, 3).reduce((sum, emp) => 
                    sum + calculateTotalScore(emp), 0) / Math.min(3, managerGroups[a].length);
                    
                topScoreB = managerGroups[b].slice(0, 3).reduce((sum, emp) => 
                    sum + calculateTotalScore(emp), 0) / Math.min(3, managerGroups[b].length);
            } else {
                // For specific metrics, use only that metric
                topScoreA = managerGroups[a].slice(0, 3).reduce((sum, emp) => 
                    sum + (parseFloat(emp[selectedMetric]) || 0), 0) / Math.min(3, managerGroups[a].length);
                    
                topScoreB = managerGroups[b].slice(0, 3).reduce((sum, emp) => 
                    sum + (parseFloat(emp[selectedMetric]) || 0), 0) / Math.min(3, managerGroups[b].length);
            }
                
            return topScoreB - topScoreA;
        });
        
        // Find top manager
        const topManager = sortedManagers[0];
        if (topManager) {
            const topEmployee = managerGroups[topManager][0];
            if (selectedMetric === 'all') {
                topManagerInfo.innerHTML = `
                    <div>🏆 ${topManager}</div>
                    <div>Top Employee: ${topEmployee.Name} (Total Score: ${calculateTotalScore(topEmployee).toFixed(1)})</div>
                `;
            } else {
                topManagerInfo.innerHTML = `
                    <div>🏆 ${topManager}</div>
                    <div>Top Employee: ${topEmployee.Name} (${metricLabels[selectedMetric]}: ${topEmployee[selectedMetric]})</div>
                `;
            }
        }
        
        // Render manager cards
        managerCards.innerHTML = '';
        sortedManagers.forEach((manager, index) => {
            const isTopManager = index === 0 && !selectedManager;
            const topEmployees = managerGroups[manager].slice(0, 5);
            
            // Calculate average score based on selected metric
            let avgScore;
            if (selectedMetric === 'all') {
                avgScore = (topEmployees.reduce((sum, emp) => 
                    sum + calculateTotalScore(emp), 0) / topEmployees.length).toFixed(1);
            } else {
                avgScore = (topEmployees.reduce((sum, emp) => 
                    sum + (parseFloat(emp[selectedMetric]) || 0), 0) / topEmployees.length).toFixed(1);
            }
            
            const card = document.createElement('div');
            card.className = `manager-card ${isTopManager ? 'highlight' : ''}`;
            
            card.innerHTML = `
                <div class="manager-header">
                    <div class="manager-name">${manager}</div>
                    <div class="manager-score">Avg: ${avgScore} ${selectedMetric === 'all' ? '(Total)' : ''}</div>
                </div>
                <div class="employee-list">
                    ${topEmployees.map(emp => {
                        const score = selectedMetric === 'all' 
                            ? calculateTotalScore(emp).toFixed(1)
                            : (emp[selectedMetric] || 0);
                            
                        return `
                            <div class="employee-item">
                                <span class="employee-name">${emp.Name}</span>
                                <span class="employee-score">${score}</span>
                            </div>
                        `;
                    }).join('')}
                </div>
            `;
            
            managerCards.appendChild(card);
        });
    }
    
    // Initialize with sample data if needed
    function initializeWithSampleData() {
        // This is just for demonstration
        // In a real app, you would load data from a file
        const sampleData = [
            { Name: 'John Doe', Manager: 'Sarah Smith', mobileSales: 45, fiberSales: 32, promoter: 78, videoPackage: 56 },
            { Name: 'Jane Smith', Manager: 'Mike Johnson', mobileSales: 67, fiberSales: 41, promoter: 89, videoPackage: 62 },
            { Name: 'Robert Brown', Manager: 'Sarah Smith', mobileSales: 52, fiberSales: 38, promoter: 65, videoPackage: 71 },
            // Add more sample data as needed
        ];
        
        processEmployeeData(sampleData);
        updateDashboard();
    }
    
    // Uncomment to use sample data
    // initializeWithSampleData();
});
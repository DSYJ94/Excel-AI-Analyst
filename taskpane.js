/* Excel AI Assistant - Web Hosted JavaScript */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('createTemplateBtn').onclick = createTemplate;
        document.getElementById('analyzeDataBtn').onclick = analyzeData;
        document.getElementById('generateFormulaBtn').onclick = generateFormula;
        document.getElementById('buildDashboardBtn').onclick = buildDashboard;
        document.getElementById('sendButton').onclick = sendMessage;
        document.getElementById('messageInput').onkeydown = handleKeyPress;
        document.getElementById('aiProvider').onchange = handleProviderChange;
        document.getElementById('saveApiKey').onclick = saveApiKey;
        
        // Load saved settings
        loadSettings();
        
        // Update status
        updateStatus('Ready', 'success');
        
        console.log('Excel AI Assistant loaded successfully');
    }
});

// Global variables
let currentProvider = 'simulation';
let apiKey = '';

// Excel Automation Functions
async function createTemplate() {
    try {
        updateStatus('Creating template...', 'processing');
        addChatMessage('user', 'Create a comprehensive budget template');
        
        await Excel.run(async (context) => {
            // Create new worksheet
            const worksheets = context.workbook.worksheets;
            const newWorksheet = worksheets.add('Budget_Template');
            
            // Set up the template structure
            const range = newWorksheet.getRange('A1:D20');
            
            // Header
            newWorksheet.getRange('A1:D1').merge();
            newWorksheet.getRange('A1').values = [['Monthly Budget Template']];
            newWorksheet.getRange('A1').format.font.bold = true;
            newWorksheet.getRange('A1').format.font.size = 16;
            newWorksheet.getRange('A1').format.horizontalAlignment = 'Center';
            newWorksheet.getRange('A1:D1').format.fill.color = '#4472C4';
            newWorksheet.getRange('A1:D1').format.font.color = 'white';
            
            // Income section
            newWorksheet.getRange('A3').values = [['INCOME']];
            newWorksheet.getRange('A3').format.font.bold = true;
            newWorksheet.getRange('A3').format.font.size = 14;
            newWorksheet.getRange('A3:D3').format.fill.color = '#E2EFDA';
            
            const incomeItems = [
                ['Salary', '', 0],
                ['Freelance Income', '', 0],
                ['Investment Returns', '', 0],
                ['Other Income', '', 0]
            ];
            
            newWorksheet.getRange('A4:C7').values = incomeItems;
            
            // Total Income
            newWorksheet.getRange('A8').values = [['Total Income']];
            newWorksheet.getRange('A8').format.font.bold = true;
            newWorksheet.getRange('C8').formulas = [['=SUM(C4:C7)']];
            newWorksheet.getRange('C8').format.font.bold = true;
            
            // Expenses section
            newWorksheet.getRange('A10').values = [['EXPENSES']];
            newWorksheet.getRange('A10').format.font.bold = true;
            newWorksheet.getRange('A10').format.font.size = 14;
            newWorksheet.getRange('A10:D10').format.fill.color = '#FCE4D6';
            
            const expenseItems = [
                ['Rent/Mortgage', '', 0],
                ['Utilities', '', 0],
                ['Food & Groceries', '', 0],
                ['Transportation', '', 0],
                ['Insurance', '', 0],
                ['Entertainment', '', 0],
                ['Savings', '', 0],
                ['Other Expenses', '', 0]
            ];
            
            newWorksheet.getRange('A11:C18').values = expenseItems;
            
            // Total Expenses
            newWorksheet.getRange('A19').values = [['Total Expenses']];
            newWorksheet.getRange('A19').format.font.bold = true;
            newWorksheet.getRange('C19').formulas = [['=SUM(C11:C18)']];
            newWorksheet.getRange('C19').format.font.bold = true;
            
            // Net Income
            newWorksheet.getRange('A21').values = [['Net Income']];
            newWorksheet.getRange('A21').format.font.bold = true;
            newWorksheet.getRange('A21').format.font.size = 14;
            newWorksheet.getRange('C21').formulas = [['=C8-C19']];
            newWorksheet.getRange('C21').format.font.bold = true;
            newWorksheet.getRange('A21:C21').format.fill.color = '#D5E8D4';
            
            // Format currency
            newWorksheet.getRange('C4:C21').numberFormat = [['$#,##0.00']];
            
            // Add borders
            newWorksheet.getRange('A1:C21').format.borders.getItem('InsideHorizontal').style = 'Continuous';
            newWorksheet.getRange('A1:C21').format.borders.getItem('InsideVertical').style = 'Continuous';
            newWorksheet.getRange('A1:C21').format.borders.getItem('EdgeBottom').style = 'Continuous';
            newWorksheet.getRange('A1:C21').format.borders.getItem('EdgeLeft').style = 'Continuous';
            newWorksheet.getRange('A1:C21').format.borders.getItem('EdgeRight').style = 'Continuous';
            newWorksheet.getRange('A1:C21').format.borders.getItem('EdgeTop').style = 'Continuous';
            
            // Auto-fit columns
            newWorksheet.getRange('A:C').format.autofitColumns();
            
            // Activate the new worksheet
            newWorksheet.activate();
            
            await context.sync();
        });
        
        const response = await getAIResponse('Create budget template', currentProvider);
        addChatMessage('ai', response);
        updateStatus('Template created successfully!', 'success');
        
    } catch (error) {
        console.error('Error creating template:', error);
        addChatMessage('ai', `Error creating template: ${error.message}`);
        updateStatus('Error creating template', 'error');
    }
}

async function analyzeData() {
    try {
        updateStatus('Analyzing data...', 'processing');
        addChatMessage('user', 'Analyze the selected data');
        
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load('values, rowCount, columnCount, address');
            
            await context.sync();
            
            if (selectedRange.rowCount < 2 || selectedRange.columnCount < 1) {
                throw new Error('Please select a range with at least 2 rows of data');
            }
            
            // Create analysis worksheet
            const worksheets = context.workbook.worksheets;
            const analysisSheet = worksheets.add('Data_Analysis');
            
            // Analysis header
            analysisSheet.getRange('A1').values = [['Data Analysis Results']];
            analysisSheet.getRange('A1').format.font.bold = true;
            analysisSheet.getRange('A1').format.font.size = 16;
            analysisSheet.getRange('A1:E1').format.fill.color = '#4472C4';
            analysisSheet.getRange('A1:E1').format.font.color = 'white';
            
            // Source information
            analysisSheet.getRange('A3').values = [['Source Range:']];
            analysisSheet.getRange('B3').values = [[selectedRange.address]];
            analysisSheet.getRange('A4').values = [['Data Points:']];
            analysisSheet.getRange('B4').values = [[selectedRange.rowCount * selectedRange.columnCount]];
            
            // Statistical analysis
            analysisSheet.getRange('A6').values = [['Statistical Summary']];
            analysisSheet.getRange('A6').format.font.bold = true;
            analysisSheet.getRange('A6').format.font.size = 14;
            
            // Calculate statistics using Excel functions
            const dataRange = selectedRange.address;
            
            analysisSheet.getRange('A8').values = [['Count:']];
            analysisSheet.getRange('B8').formulas = [[`=COUNTA(${dataRange})`]];
            
            analysisSheet.getRange('A9').values = [['Sum:']];
            analysisSheet.getRange('B9').formulas = [[`=SUM(${dataRange})`]];
            
            analysisSheet.getRange('A10').values = [['Average:']];
            analysisSheet.getRange('B10').formulas = [[`=AVERAGE(${dataRange})`]];
            
            analysisSheet.getRange('A11').values = [['Median:']];
            analysisSheet.getRange('B11').formulas = [[`=MEDIAN(${dataRange})`]];
            
            analysisSheet.getRange('A12').values = [['Maximum:']];
            analysisSheet.getRange('B12').formulas = [[`=MAX(${dataRange})`]];
            
            analysisSheet.getRange('A13').values = [['Minimum:']];
            analysisSheet.getRange('B13').formulas = [[`=MIN(${dataRange})`]];
            
            analysisSheet.getRange('A14').values = [['Standard Deviation:']];
            analysisSheet.getRange('B14').formulas = [[`=STDEV(${dataRange})`]];
            
            // Format numbers
            analysisSheet.getRange('B9:B14').numberFormat = [['#,##0.00']];
            
            // Add chart if data is suitable
            if (selectedRange.rowCount > 1 && selectedRange.columnCount >= 1) {
                const chart = analysisSheet.charts.add('ColumnClustered', selectedRange, 'Auto');
                chart.setPosition('D6', 'H20');
                chart.title.text = 'Data Visualization';
            }
            
            // Auto-fit columns
            analysisSheet.getRange('A:E').format.autofitColumns();
            
            // Activate analysis sheet
            analysisSheet.activate();
            
            await context.sync();
        });
        
        const response = await getAIResponse('Analyze data', currentProvider);
        addChatMessage('ai', response);
        updateStatus('Data analysis completed!', 'success');
        
    } catch (error) {
        console.error('Error analyzing data:', error);
        addChatMessage('ai', `Error analyzing data: ${error.message}`);
        updateStatus('Error analyzing data', 'error');
    }
}

async function generateFormula() {
    try {
        updateStatus('Generating formulas...', 'processing');
        addChatMessage('user', 'Generate advanced Excel formulas');
        
        await Excel.run(async (context) => {
            // Create formulas worksheet
            const worksheets = context.workbook.worksheets;
            const formulaSheet = worksheets.add('Advanced_Formulas');
            
            // Header
            formulaSheet.getRange('A1').values = [['Advanced Excel Formulas']];
            formulaSheet.getRange('A1').format.font.bold = true;
            formulaSheet.getRange('A1').format.font.size = 16;
            formulaSheet.getRange('A1:C1').format.fill.color = '#4472C4';
            formulaSheet.getRange('A1:C1').format.font.color = 'white';
            
            // Formula examples
            const formulas = [
                ['Formula Type', 'Example', 'Description'],
                ['VLOOKUP', '=VLOOKUP(A2,D:F,2,FALSE)', 'Lookup value in table'],
                ['INDEX/MATCH', '=INDEX(F:F,MATCH(A2,D:D,0))', 'Flexible lookup function'],
                ['SUMIFS', '=SUMIFS(C:C,A:A,"Criteria1",B:B,">100")', 'Sum with multiple criteria'],
                ['COUNTIFS', '=COUNTIFS(A:A,"Product",B:B,">50")', 'Count with multiple criteria'],
                ['IF with AND', '=IF(AND(A2>100,B2<50),"Good","Review")', 'Multiple condition check'],
                ['IF with OR', '=IF(OR(A2="A",A2="B"),"Priority","Normal")', 'Alternative condition check'],
                ['IFERROR', '=IFERROR(VLOOKUP(A2,D:F,2,FALSE),"Not Found")', 'Error handling'],
                ['CONCATENATE', '=CONCATENATE(A2," - ",B2)', 'Combine text values'],
                ['TEXT', '=TEXT(A2,"mm/dd/yyyy")', 'Format numbers as text'],
                ['ROUND', '=ROUND(A2*B2,2)', 'Round to specific decimals'],
                ['NETWORKDAYS', '=NETWORKDAYS(A2,B2)', 'Business days between dates'],
                ['SUMPRODUCT', '=SUMPRODUCT((A:A="Criteria")*(B:B))', 'Array-like calculations'],
                ['CHOOSE', '=CHOOSE(A2,"Option1","Option2","Option3")', 'Select from list by index']
            ];
            
            formulaSheet.getRange('A3:C16').values = formulas;
            
            // Format header row
            formulaSheet.getRange('A3:C3').format.font.bold = true;
            formulaSheet.getRange('A3:C3').format.fill.color = '#E2EFDA';
            
            // Add borders
            formulaSheet.getRange('A3:C16').format.borders.getItem('InsideHorizontal').style = 'Continuous';
            formulaSheet.getRange('A3:C16').format.borders.getItem('InsideVertical').style = 'Continuous';
            formulaSheet.getRange('A3:C16').format.borders.getItem('EdgeBottom').style = 'Continuous';
            formulaSheet.getRange('A3:C16').format.borders.getItem('EdgeLeft').style = 'Continuous';
            formulaSheet.getRange('A3:C16').format.borders.getItem('EdgeRight').style = 'Continuous';
            formulaSheet.getRange('A3:C16').format.borders.getItem('EdgeTop').style = 'Continuous';
            
            // Auto-fit columns
            formulaSheet.getRange('A:C').format.autofitColumns();
            
            // Activate formula sheet
            formulaSheet.activate();
            
            await context.sync();
        });
        
        const response = await getAIResponse('Generate formulas', currentProvider);
        addChatMessage('ai', response);
        updateStatus('Formulas generated successfully!', 'success');
        
    } catch (error) {
        console.error('Error generating formulas:', error);
        addChatMessage('ai', `Error generating formulas: ${error.message}`);
        updateStatus('Error generating formulas', 'error');
    }
}

async function buildDashboard() {
    try {
        updateStatus('Building dashboard...', 'processing');
        addChatMessage('user', 'Build an executive dashboard');
        
        await Excel.run(async (context) => {
            // Create dashboard worksheet
            const worksheets = context.workbook.worksheets;
            const dashboardSheet = worksheets.add('Executive_Dashboard');
            
            // Dashboard title
            dashboardSheet.getRange('A1:F1').merge();
            dashboardSheet.getRange('A1').values = [['Executive Dashboard']];
            dashboardSheet.getRange('A1').format.font.bold = true;
            dashboardSheet.getRange('A1').format.font.size = 20;
            dashboardSheet.getRange('A1').format.horizontalAlignment = 'Center';
            dashboardSheet.getRange('A1:F1').format.fill.color = '#1F4E79';
            dashboardSheet.getRange('A1:F1').format.font.color = 'white';
            
            // KPI Section
            dashboardSheet.getRange('A3').values = [['Key Performance Indicators']];
            dashboardSheet.getRange('A3').format.font.bold = true;
            dashboardSheet.getRange('A3').format.font.size = 14;
            
            // KPI Cards
            const kpis = [
                ['Revenue', '$1,250,000', '↗ 15.3%'],
                ['Profit Margin', '23.7%', '↗ 2.1%'],
                ['Customer Satisfaction', '94.2%', '↗ 1.8%'],
                ['Market Share', '18.5%', '↗ 0.7%']
            ];
            
            // Create KPI cards
            for (let i = 0; i < kpis.length; i++) {
                const row = 5 + (i * 3);
                const kpi = kpis[i];
                
                // KPI Name
                dashboardSheet.getRange(`A${row}`).values = [[kpi[0]]];
                dashboardSheet.getRange(`A${row}`).format.font.bold = true;
                dashboardSheet.getRange(`A${row}`).format.font.size = 12;
                
                // KPI Value
                dashboardSheet.getRange(`B${row}`).values = [[kpi[1]]];
                dashboardSheet.getRange(`B${row}`).format.font.bold = true;
                dashboardSheet.getRange(`B${row}`).format.font.size = 16;
                dashboardSheet.getRange(`B${row}`).format.font.color = '#1F4E79';
                
                // KPI Change
                dashboardSheet.getRange(`C${row}`).values = [[kpi[2]]];
                dashboardSheet.getRange(`C${row}`).format.font.color = '#0F7B0F';
                dashboardSheet.getRange(`C${row}`).format.font.bold = true;
                
                // Add border around KPI card
                dashboardSheet.getRange(`A${row}:C${row + 1}`).format.borders.getItem('EdgeBottom').style = 'Continuous';
                dashboardSheet.getRange(`A${row}:C${row + 1}`).format.borders.getItem('EdgeLeft').style = 'Continuous';
                dashboardSheet.getRange(`A${row}:C${row + 1}`).format.borders.getItem('EdgeRight').style = 'Continuous';
                dashboardSheet.getRange(`A${row}:C${row + 1}`).format.borders.getItem('EdgeTop').style = 'Continuous';
                dashboardSheet.getRange(`A${row}:C${row + 1}`).format.fill.color = '#F2F2F2';
            }
            
            // Sample data for charts
            dashboardSheet.getRange('E3').values = [['Monthly Performance']];
            dashboardSheet.getRange('E3').format.font.bold = true;
            dashboardSheet.getRange('E3').format.font.size = 14;
            
            const monthlyData = [
                ['Month', 'Revenue', 'Profit'],
                ['Jan', 980000, 230000],
                ['Feb', 1050000, 245000],
                ['Mar', 1120000, 265000],
                ['Apr', 1180000, 280000],
                ['May', 1250000, 295000],
                ['Jun', 1320000, 315000]
            ];
            
            dashboardSheet.getRange('E5:G11').values = monthlyData;
            
            // Format data table
            dashboardSheet.getRange('E5:G5').format.font.bold = true;
            dashboardSheet.getRange('E5:G5').format.fill.color = '#E2EFDA';
            dashboardSheet.getRange('F6:G11').numberFormat = [['$#,##0']];
            
            // Add chart
            const chartRange = dashboardSheet.getRange('E5:G11');
            const chart = dashboardSheet.charts.add('ColumnClustered', chartRange, 'Auto');
            chart.setPosition('E13', 'K25');
            chart.title.text = 'Revenue & Profit Trend';
            chart.legend.position = 'Bottom';
            
            // Auto-fit columns
            dashboardSheet.getRange('A:K').format.autofitColumns();
            
            // Activate dashboard
            dashboardSheet.activate();
            
            await context.sync();
        });
        
        const response = await getAIResponse('Build dashboard', currentProvider);
        addChatMessage('ai', response);
        updateStatus('Dashboard created successfully!', 'success');
        
    } catch (error) {
        console.error('Error building dashboard:', error);
        addChatMessage('ai', `Error building dashboard: ${error.message}`);
        updateStatus('Error building dashboard', 'error');
    }
}

// Chat Functions
async function sendMessage() {
    const messageInput = document.getElementById('messageInput');
    const message = messageInput.value.trim();
    
    if (!message) return;
    
    addChatMessage('user', message);
    messageInput.value = '';
    
    updateStatus('Processing...', 'processing');
    
    try {
        // Process the command
        await processCommand(message);
        
        // Get AI response
        const response = await getAIResponse(message, currentProvider);
        addChatMessage('ai', response);
        
        updateStatus('Ready', 'success');
    } catch (error) {
        console.error('Error processing message:', error);
        addChatMessage('ai', `Error: ${error.message}`);
        updateStatus('Error', 'error');
    }
}

async function processCommand(command) {
    const lowerCommand = command.toLowerCase();
    
    if (lowerCommand.includes('template') || lowerCommand.includes('budget')) {
        await createTemplate();
    } else if (lowerCommand.includes('analyze') || lowerCommand.includes('data')) {
        await analyzeData();
    } else if (lowerCommand.includes('formula') || lowerCommand.includes('calculation')) {
        await generateFormula();
    } else if (lowerCommand.includes('dashboard') || lowerCommand.includes('chart')) {
        await buildDashboard();
    }
}

function handleKeyPress(event) {
    if (event.ctrlKey && event.key === 'Enter') {
        event.preventDefault();
        sendMessage();
    }
}

// AI Service Functions
async function getAIResponse(message, provider) {
    if (provider === 'simulation') {
        return simulateAIResponse(message);
    }
    
    // For real AI providers, implement API calls here
    try {
        const response = await callAIProvider(message, provider);
        return response;
    } catch (error) {
        return `Error connecting to ${provider}: ${error.message}`;
    }
}

function simulateAIResponse(message) {
    const responses = {
        'template': '✅ Budget Template Created!\n📊 Features Added:\n• Income categories with automatic formulas\n• Expense tracking with calculations\n• Net income calculation with conditional formatting\n• Professional color coding and borders\n• Currency formatting throughout\n• Created in new worksheet: Budget_Template',
        
        'analyze': '🔍 Data Analysis Complete!\n📈 Analysis Results:\n• Statistical summary generated\n• Count, sum, average, median calculated\n• Maximum and minimum values identified\n• Standard deviation computed\n• Visual chart created for data trends\n• Results saved in: Data_Analysis worksheet',
        
        'formula': '⚡ Advanced Formulas Generated!\n🧮 Formula Library Created:\n• VLOOKUP for data lookup operations\n• INDEX/MATCH for flexible searches\n• SUMIFS/COUNTIFS for conditional calculations\n• IF statements with AND/OR logic\n• Error handling with IFERROR\n• Date and text manipulation functions\n• Created in: Advanced_Formulas worksheet',
        
        'dashboard': '📈 Executive Dashboard Built!\n🎯 Dashboard Components:\n• 4 Key Performance Indicators (KPIs)\n• Revenue and profit tracking\n• Performance trend visualization\n• Color-coded status indicators\n• Professional executive-level formatting\n• Interactive charts and data tables\n• Created in: Executive_Dashboard worksheet'
    };
    
    for (const key in responses) {
        if (message.toLowerCase().includes(key)) {
            return responses[key];
        }
    }
    
    return `I understand you want to: "${message}"\n\nI can help you with:\n• Creating comprehensive templates (budgets, financial models)\n• Analyzing data with statistical insights\n• Generating advanced Excel formulas\n• Building interactive dashboards\n\nTry using the quick action buttons or type a more specific command!`;
}

async function callAIProvider(message, provider) {
    // This would implement actual API calls to AI providers
    // For now, return a placeholder
    return `This would call ${provider} API with message: ${message}`;
}

// UI Helper Functions
function addChatMessage(sender, message) {
    const chatMessages = document.getElementById('chatMessages');
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${sender}-message`;
    
    const avatar = document.createElement('div');
    avatar.className = 'message-avatar';
    avatar.textContent = sender === 'user' ? '👤' : '🤖';
    
    const content = document.createElement('div');
    content.className = 'message-content';
    
    const text = document.createElement('div');
    text.className = 'message-text';
    text.innerHTML = message.replace(/\n/g, '<br>');
    
    const time = document.createElement('div');
    time.className = 'message-time';
    time.textContent = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    
    content.appendChild(text);
    content.appendChild(time);
    messageDiv.appendChild(avatar);
    messageDiv.appendChild(content);
    
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

function updateStatus(text, type) {
    const statusIndicator = document.getElementById('statusIndicator');
    const statusText = statusIndicator.querySelector('.status-text');
    
    statusText.textContent = text;
    statusIndicator.className = `status-indicator ${type}`;
}

function handleProviderChange() {
    const provider = document.getElementById('aiProvider').value;
    const apiKeyGroup = document.getElementById('apiKeyGroup');
    
    currentProvider = provider;
    
    if (provider === 'simulation') {
        apiKeyGroup.style.display = 'none';
    } else {
        apiKeyGroup.style.display = 'block';
    }
    
    saveSettings();
}

function saveApiKey() {
    const apiKeyInput = document.getElementById('apiKey');
    apiKey = apiKeyInput.value;
    
    if (apiKey) {
        localStorage.setItem('excelai_apikey', apiKey);
        updateStatus('API key saved', 'success');
        setTimeout(() => updateStatus('Ready', 'success'), 2000);
    }
}

function loadSettings() {
    const savedProvider = localStorage.getItem('excelai_provider');
    const savedApiKey = localStorage.getItem('excelai_apikey');
    
    if (savedProvider) {
        document.getElementById('aiProvider').value = savedProvider;
        currentProvider = savedProvider;
        handleProviderChange();
    }
    
    if (savedApiKey) {
        document.getElementById('apiKey').value = savedApiKey;
        apiKey = savedApiKey;
    }
}

function saveSettings() {
    localStorage.setItem('excelai_provider', currentProvider);
}


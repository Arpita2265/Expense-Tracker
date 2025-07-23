
        document.addEventListener('DOMContentLoaded', function() {
            // DOM Elements
            const menuToggle = document.getElementById('menuToggle');
            const sidebar = document.getElementById('sidebar');
            const sidebarOverlay = document.getElementById('sidebarOverlay');
            const closeSidebar = document.getElementById('closeSidebar');
            const mainContent = document.getElementById('mainContent');
            const expenseForm = document.getElementById('expenseForm');
            const expenseTableBody = document.getElementById('expenseTableBody');
            const downloadExcelBtn = document.getElementById('downloadExcel');
            const saveFileBtn = document.getElementById('saveFileBtn');
            const newFileBtn = document.getElementById('newFileBtn');
            const saveFileModal = document.getElementById('saveFileModal');
            const closeSaveModal = document.getElementById('closeSaveModal');
            const cancelSave = document.getElementById('cancelSave');
            const confirmSave = document.getElementById('confirmSave');
            const fileNameInput = document.getElementById('fileName');
            const homeLink = document.getElementById('homeLink');
            const historyLink = document.getElementById('historyLink');
            const filesLink = document.getElementById('filesLink');
            const customLink = document.getElementById('customLink');
            const homePage = document.getElementById('homePage');
            const historyPage = document.getElementById('historyPage');
            const filesPage = document.getElementById('filesPage');
            const customPage = document.getElementById('customPage');
            const historyList = document.getElementById('historyList');
            const filesTableBody = document.getElementById('filesTableBody');
            const noFilesMessage = document.getElementById('noFilesMessage');
            const editFileModal = document.getElementById('editFileModal');
            const closeEditModal = document.getElementById('closeEditModal');
            const cancelEdit = document.getElementById('cancelEdit');
            const confirmEdit = document.getElementById('confirmEdit');
            const editFileNameInput = document.getElementById('editFileName');
            const stayYesRadio = document.getElementById('stayYes');
            const stayNoRadio = document.getElementById('stayNo');
            const stayChargeGroup = document.getElementById('stayChargeGroup');
            const themeToggle = document.getElementById('themeToggle');
            const sidebarSearch = document.getElementById('sidebarSearch');
            const filesSearch = document.getElementById('filesSearch');
            
            // Custom Page Elements
            const columnNameInput = document.getElementById('columnName');
            const columnTypeSelect = document.getElementById('columnType');
            const optionsGroup = document.getElementById('optionsGroup');
            const columnOptionsInput = document.getElementById('columnOptions');
            const addColumnBtn = document.getElementById('addColumnBtn');
            const columnList = document.getElementById('columnList');
            const customTableNameInput = document.getElementById('customTableName');
            const createCustomTableBtn = document.getElementById('createCustomTableBtn');
            const customTableContainer = document.getElementById('customTableContainer');
            const customTableTitle = document.getElementById('customTableTitle');
            const customTableForm = document.getElementById('customTableForm');
            const customTableHead = document.getElementById('customTableHead');
            const customTableBody = document.getElementById('customTableBody');
            const downloadCustomExcelBtn = document.getElementById('downloadCustomExcel');
            const saveCustomFileBtn = document.getElementById('saveCustomFileBtn');

            // Data
            let expenses = [];
            let currentFile = {
                id: Date.now(),
                name: 'Untitled',
                data: [],
                createdAt: new Date(),
                type: 'expense'
            };
            let files = JSON.parse(localStorage.getItem('expenseFiles')) || [];
            let history = JSON.parse(localStorage.getItem('expenseHistory')) || [];
            let fileToEdit = null;
            
            // Custom Table Data
            let customColumns = [];
            let customTableData = [];
            let currentCustomFile = {
                id: Date.now(),
                name: 'Untitled Custom',
                data: [],
                columns: [],
                createdAt: new Date(),
                type: 'custom'
            };

            // Initialize the app
            function init() {
                // Check for saved theme preference
                if (localStorage.getItem('darkMode') === 'enabled') {
                    document.body.classList.add('dark-mode');
                    themeToggle.checked = true;
                }
                
                // Load the most recent file if available
                if (history.length > 0) {
                    const lastFileId = history[0].id;
                    const lastFile = files.find(file => file.id === lastFileId);
                    if (lastFile) {
                        loadFile(lastFile.id, false);
                    }
                }
                
                // Load pages
                loadPage('home');
            }

            // Theme toggle
            themeToggle.addEventListener('change', function() {
                if (this.checked) {
                    document.body.classList.add('dark-mode');
                    localStorage.setItem('darkMode', 'enabled');
                } else {
                    document.body.classList.remove('dark-mode');
                    localStorage.setItem('darkMode', 'disabled');
                }
            });

            // Sidebar toggle
            menuToggle.addEventListener('click', function() {
                sidebar.classList.add('open');
                sidebarOverlay.classList.add('open');
                mainContent.classList.add('expanded');
            });

            closeSidebar.addEventListener('click', function() {
                sidebar.classList.remove('open');
                sidebarOverlay.classList.remove('open');
                mainContent.classList.remove('expanded');
            });

            sidebarOverlay.addEventListener('click', function() {
                sidebar.classList.remove('open');
                sidebarOverlay.classList.remove('open');
                mainContent.classList.remove('expanded');
            });

            // Navigation
            homeLink.addEventListener('click', function() {
                loadPage('home');
                sidebar.classList.remove('open');
                sidebarOverlay.classList.remove('open');
                mainContent.classList.remove('expanded');
            });

            historyLink.addEventListener('click', function() {
                loadPage('history');
                renderHistoryList();
                sidebar.classList.remove('open');
                sidebarOverlay.classList.remove('open');
                mainContent.classList.remove('expanded');
            });

            filesLink.addEventListener('click', function() {
                loadPage('files');
                renderFilesList();
                sidebar.classList.remove('open');
                sidebarOverlay.classList.remove('open');
                mainContent.classList.remove('expanded');
            });

            customLink.addEventListener('click', function() {
                loadPage('custom');
                sidebar.classList.remove('open');
                sidebarOverlay.classList.remove('open');
                mainContent.classList.remove('expanded');
            });

            function loadPage(page) {
                // Hide all pages
                document.querySelectorAll('.page').forEach(p => {
                    p.classList.remove('active');
                });
                
                // Show the selected page
                switch(page) {
                    case 'home':
                        homePage.classList.add('active');
                        document.title = 'Expense Tracker - Home';
                        break;
                    case 'history':
                        historyPage.classList.add('active');
                        document.title = 'Expense Tracker - History';
                        break;
                    case 'files':
                        filesPage.classList.add('active');
                        document.title = 'Expense Tracker - Files';
                        break;
                    case 'custom':
                        customPage.classList.add('active');
                        document.title = 'Expense Tracker - Custom';
                        break;
                }
            }

            // Stay charge visibility
            stayYesRadio.addEventListener('change', function() {
                if (this.checked) {
                    stayChargeGroup.style.display = 'block';
                }
            });

            stayNoRadio.addEventListener('change', function() {
                if (this.checked) {
                    stayChargeGroup.style.display = 'none';
                }
            });

            // Form submission
            expenseForm.addEventListener('submit', function(e) {
                e.preventDefault();
                
                const date = document.getElementById('date').value;
                const from = document.getElementById('from').value;
                const to = document.getElementById('to').value;
                const travellingCost = document.getElementById('travellingCost').value;
                const ticket = document.querySelector('input[name="ticket"]:checked').value;
                const travellingClaim = document.getElementById('travellingClaim').value;
                const ticketClaim = document.querySelector('input[name="ticketClaim"]:checked').value;
                const da = document.getElementById('da').value;
                const daReceipt = document.querySelector('input[name="daReceipt"]:checked').value;
                const daClaim = document.getElementById('daClaim').value;
                const stay = document.querySelector('input[name="stay"]:checked').value;
                const stayCharge = stay === 'Yes' ? document.getElementById('stayCharge').value : 0;
                const daClaimReceipt = document.querySelector('input[name="daClaimReceipt"]:checked').value;
                
                // Format the data for the table
                const destination = `${from} ➝ ${to}`;
                
                let formattedTravellingCost = travellingCost;
                formattedTravellingCost = `Ticket ➝ ${ticket}\n${formattedTravellingCost}`;
                
                let formattedTravellingClaim = travellingClaim;
                formattedTravellingClaim = `Ticket ➝ ${ticketClaim}\n${formattedTravellingClaim}`;
                
                // Calculate DA sum with proper format
                let daSum = 0;
                let formattedDA = da;
                if (da.includes('➝')) {
                    const daItems = da.split('+');
                    daSum = daItems.reduce((sum, item) => {
                        const parts = item.split('➝');
                        const value = parts.length > 1 ? parseFloat(parts[1].trim()) || 0 : parseFloat(parts[0].trim()) || 0;
                        return sum + value;
                    }, 0);
                    formattedDA = `${da}= ${daSum}`;
                } else {
                    daSum = calculateSum(da);
                    formattedDA = da;
                }
                
                formattedDA = `Receipt ➝ ${daReceipt}\n${formattedDA}`;
                
                let formattedDAClaim = daClaim;
                if (stay === 'Yes') {
                    formattedDAClaim = `Stay ➝ ${stayCharge}\n${formattedDAClaim}`;
                }
                formattedDAClaim = `Receipt ➝ ${daClaimReceipt}\n${formattedDAClaim}`;
                
                // Calculate sums
                const travellingCostSum = calculateSum(travellingCost);
                const travellingClaimSum = calculateSum(travellingClaim);
                let daClaimSum = calculateSum(daClaim);
                if (stay === 'Yes') {
                    daClaimSum += parseFloat(stayCharge) || 0;
                }
                const totalClaim = travellingClaimSum + daClaimSum;
                
                // Create expense object
                const expense = {
                    id: Date.now(),
                    date,
                    destination,
                    travellingCost: formattedTravellingCost,
                    travellingCostSum,
                    travellingClaim: formattedTravellingClaim,
                    travellingClaimSum,
                    da: formattedDA,
                    daSum,
                    daClaim: formattedDAClaim,
                    daClaimSum,
                    totalClaim
                };
                
                // Add to expenses array
                expenses.push(expense);
                
                // Update current file data
                currentFile.data = expenses;
                currentFile.updatedAt = new Date();
                
                // Save to localStorage
                saveCurrentFile();
                
                // Render table
                renderExpenseTable();
                updateSummary();
                
                // Reset form
                expenseForm.reset();
                stayChargeGroup.style.display = 'none';
            });

            // Calculate sum from a string like "100+50+30"
            function calculateSum(valueString) {
                if (!valueString) return 0;
                
                try {
                    return valueString.split('+').reduce((sum, val) => {
                        return sum + (parseFloat(val.trim()) || 0);
                    }, 0);
                } catch (e) {
                    console.error('Error calculating sum:', e);
                    return 0;
                }
            }

            // Render expense table
            function renderExpenseTable() {
                expenseTableBody.innerHTML = '';
                
                if (expenses.length === 0) {
                    const row = document.createElement('tr');
                    row.innerHTML = `<td colspan="8" style="text-align: center;">No expenses added yet</td>`;
                    expenseTableBody.appendChild(row);
                    
                    // Reset subtotals
                    document.getElementById('subtotalTravellingCost').textContent = '0';
                    document.getElementById('subtotalTravellingClaim').textContent = '0';
                    document.getElementById('subtotalDA').textContent = '0';
                    document.getElementById('subtotalDAClaim').textContent = '0';
                    document.getElementById('subtotalTotalClaim').textContent = '0';
                    
                    return;
                }
                
                // Group expenses by date and destination for subtotals
                const groupedExpenses = {};
                expenses.forEach(expense => {
                    const key = `${expense.date}_${expense.destination}`;
                    if (!groupedExpenses[key]) {
                        groupedExpenses[key] = [];
                    }
                    groupedExpenses[key].push(expense);
                });
                
                // Render each group with subtotal
                Object.keys(groupedExpenses).forEach((key, groupIndex) => {
                    const group = groupedExpenses[key];
                    const firstExpense = group[0];
                    
                    // Calculate subtotals for the group
                    const subtotalTravellingCost = group.reduce((sum, exp) => sum + exp.travellingCostSum, 0);
                    const subtotalTravellingClaim = group.reduce((sum, exp) => sum + exp.travellingClaimSum, 0);
                    const subtotalDA = group.reduce((sum, exp) => sum + exp.daSum, 0);
                    const subtotalDAClaim = group.reduce((sum, exp) => sum + exp.daClaimSum, 0);
                    const subtotalTotalClaim = group.reduce((sum, exp) => sum + exp.totalClaim, 0);
                    
                    // Render each expense in the group
                    group.forEach((expense, index) => {
                        const row = document.createElement('tr');
                        
                        // Only show date and destination for the first row in the group
                        const dateCell = index === 0 ? `<td rowspan="${group.length}">${expense.date}</td>` : '';
                        const destinationCell = index === 0 ? `<td rowspan="${group.length}">${expense.destination}</td>` : '';
                        
                        row.innerHTML = `
                            ${dateCell}
                            ${destinationCell}
                            <td>${formatMultiline(expense.travellingCost)}<br><strong>${expense.travellingCostSum.toFixed(2)}</strong></td>
                            <td>${formatMultiline(expense.travellingClaim)}<br><strong>${expense.travellingClaimSum.toFixed(2)}</strong></td>
                            <td>${formatMultiline(expense.da)}<br><strong>${expense.daSum.toFixed(2)}</strong></td>
                            <td>${formatMultiline(expense.daClaim)}<br><strong>${expense.daClaimSum.toFixed(2)}</strong></td>
                            <td><strong>${expense.totalClaim.toFixed(2)}</strong></td>
                            <td>
                                <div class="action-buttons">
                                    <button class="action-btn edit-btn" data-id="${expense.id}">Edit</button>
                                    <button class="action-btn delete-btn" data-id="${expense.id}">Delete</button>
                                </div>
                            </td>
                        `;
                        
                        expenseTableBody.appendChild(row);
                    });
                    
                    // Add subtotal row after each group (except the last one)
                    if (groupIndex < Object.keys(groupedExpenses).length - 1) {
                        const subtotalRow = document.createElement('tr');
                        subtotalRow.className = 'subtotal-row';
                        subtotalRow.innerHTML = `
                            <td colspan="2"><strong>Subtotal</strong></td>
                            <td><strong>${subtotalTravellingCost.toFixed(2)}</strong></td>
                            <td><strong>${subtotalTravellingClaim.toFixed(2)}</strong></td>
                            <td><strong>${subtotalDA.toFixed(2)}</strong></td>
                            <td><strong>${subtotalDAClaim.toFixed(2)}</strong></td>
                            <td><strong>${subtotalTotalClaim.toFixed(2)}</strong></td>
                            <td></td>
                        `;
                        expenseTableBody.appendChild(subtotalRow);
                    }
                });
                
                // Update the footer subtotals
                updateSubtotals();
                
                // Add event listeners to action buttons
                document.querySelectorAll('.edit-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const id = parseInt(this.getAttribute('data-id'));
                        editExpense(id);
                    });
                });
                
                document.querySelectorAll('.delete-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const id = parseInt(this.getAttribute('data-id'));
                        deleteExpense(id);
                    });
                });
            }

            // Format multiline text for display
            function formatMultiline(text) {
                if (!text) return '';
                return text.replace(/\n/g, '<br>');
            }

            // Update subtotals in table footer
            function updateSubtotals() {
                const subtotalTravellingCost = expenses.reduce((sum, exp) => sum + exp.travellingCostSum, 0);
                const subtotalTravellingClaim = expenses.reduce((sum, exp) => sum + exp.travellingClaimSum, 0);
                const subtotalDA = expenses.reduce((sum, exp) => sum + exp.daSum, 0);
                const subtotalDAClaim = expenses.reduce((sum, exp) => sum + exp.daClaimSum, 0);
                const subtotalTotalClaim = expenses.reduce((sum, exp) => sum + exp.totalClaim, 0);
                
                document.getElementById('subtotalTravellingCost').textContent = subtotalTravellingCost.toFixed(2);
                document.getElementById('subtotalTravellingClaim').textContent = subtotalTravellingClaim.toFixed(2);
                document.getElementById('subtotalDA').textContent = subtotalDA.toFixed(2);
                document.getElementById('subtotalDAClaim').textContent = subtotalDAClaim.toFixed(2);
                document.getElementById('subtotalTotalClaim').textContent = subtotalTotalClaim.toFixed(2);
            }

            // Update summary section
            function updateSummary() {
                if (expenses.length === 0) {
                    document.getElementById('totalDays').textContent = '0';
                    document.getElementById('totalTravellingCost').textContent = '0';
                    document.getElementById('totalTravellingClaim').textContent = '0';
                    document.getElementById('totalDA').textContent = '0';
                    document.getElementById('totalDAClaim').textContent = '0';
                    document.getElementById('totalClaim').textContent = '0';
                    return;
                }
                
                // Calculate unique days
                const uniqueDays = new Set(expenses.map(exp => exp.date));
                document.getElementById('totalDays').textContent = uniqueDays.size;
                
                // Calculate totals
                const totalTravellingCost = expenses.reduce((sum, exp) => sum + exp.travellingCostSum, 0);
                const totalTravellingClaim = expenses.reduce((sum, exp) => sum + exp.travellingClaimSum, 0);
                const totalDA = expenses.reduce((sum, exp) => sum + exp.daSum, 0);
                const totalDAClaim = expenses.reduce((sum, exp) => sum + exp.daClaimSum, 0);
                const totalClaim = expenses.reduce((sum, exp) => sum + exp.totalClaim, 0);
                
                document.getElementById('totalTravellingCost').textContent = totalTravellingCost.toFixed(2);
                document.getElementById('totalTravellingClaim').textContent = totalTravellingClaim.toFixed(2);
                document.getElementById('totalDA').textContent = totalDA.toFixed(2);
                document.getElementById('totalDAClaim').textContent = totalDAClaim.toFixed(2);
                document.getElementById('totalClaim').textContent = totalClaim.toFixed(2);
            }

            // Edit expense
            function editExpense(id) {
                const expense = expenses.find(exp => exp.id === id);
                if (!expense) return;
                
                // Extract original values from the formatted strings
                const destinationParts = expense.destination.split(' ➝ ');
                const from = destinationParts[0];
                const to = destinationParts[1] || '';
                
                // Extract ticket availability from travellingCost
                let ticket = 'Not Available';
                if (expense.travellingCost.includes('Ticket ➝ Available')) {
                    ticket = 'Available';
                }
                let travellingCost = expense.travellingCost.replace(`Ticket ➝ ${ticket}\n`, '');
                
                // Extract ticket claim availability from travellingClaim
                let ticketClaim = 'Not Available';
                if (expense.travellingClaim.includes('Ticket ➝ Available')) {
                    ticketClaim = 'Available';
                }
                let travellingClaim = expense.travellingClaim.replace(`Ticket ➝ ${ticketClaim}\n`, '');
                
                // Extract DA receipt availability
                let daReceipt = 'Not Available';
                if (expense.da.includes('Receipt ➝ Available')) {
                    daReceipt = 'Available';
                }
                let da = expense.da.split('=')[0].trim(); // Remove the sum part
                da = da.replace(`Receipt ➝ ${daReceipt}\n`, '');
                
                // Extract DA claim receipt and stay information
                let daClaimReceipt = 'Not Available';
                let stay = 'No';
                let stayCharge = '0';
                let daClaim = expense.daClaim;
                
                if (daClaim.includes('Receipt ➝ Available')) {
                    daClaimReceipt = 'Available';
                    daClaim = daClaim.replace(`Receipt ➝ ${daClaimReceipt}\n`, '');
                } else if (daClaim.includes('Receipt ➝ Not Available')) {
                    daClaim = daClaim.replace(`Receipt ➝ Not Available\n`, '');
                }
                
                if (daClaim.includes('Stay ➝')) {
                    stay = 'Yes';
                    const stayMatch = daClaim.match(/Stay ➝ (\d+(\.\d+)?)/);
                    if (stayMatch) {
                        stayCharge = stayMatch[1];
                    }
                    daClaim = daClaim.replace(`Stay ➝ ${stayCharge}\n`, '');
                }
                
                // Fill the form with the expense data
                document.getElementById('date').value = expense.date;
                document.getElementById('from').value = from;
                document.getElementById('to').value = to;
                document.getElementById('travellingCost').value = travellingCost;
                document.querySelector(`input[name="ticket"][value="${ticket}"]`).checked = true;
                document.getElementById('travellingClaim').value = travellingClaim;
                document.querySelector(`input[name="ticketClaim"][value="${ticketClaim}"]`).checked = true;
                document.getElementById('da').value = da;
                document.querySelector(`input[name="daReceipt"][value="${daReceipt}"]`).checked = true;
                document.getElementById('daClaim').value = daClaim;
                document.querySelector(`input[name="stay"][value="${stay}"]`).checked = true;
                
                if (stay === 'Yes') {
                    stayChargeGroup.style.display = 'block';
                    document.getElementById('stayCharge').value = stayCharge;
                } else {
                    stayChargeGroup.style.display = 'none';
                }
                
                document.querySelector(`input[name="daClaimReceipt"][value="${daClaimReceipt}"]`).checked = true;
                
                // Remove the expense from the array
                expenses = expenses.filter(exp => exp.id !== id);
                
                // Update current file data
                currentFile.data = expenses;
                currentFile.updatedAt = new Date();
                
                // Save to localStorage
                saveCurrentFile();
                
                // Re-render the table
                renderExpenseTable();
                updateSummary();
            }

            // Delete expense
            function deleteExpense(id) {
                if (confirm('Are you sure you want to delete this expense?')) {
                    expenses = expenses.filter(exp => exp.id !== id);
                    
                    // Update current file data
                    currentFile.data = expenses;
                    currentFile.updatedAt = new Date();
                    
                    // Save to localStorage
                    saveCurrentFile();
                    
                    // Re-render the table
                    renderExpenseTable();
                    updateSummary();
                }
            }

            // Download Excel
            downloadExcelBtn.addEventListener('click', function() {
                if (expenses.length === 0) {
                    alert('No data to export');
                    return;
                }
                
                // Prepare data for Excel
                const excelData = [
                    ['Date', 'Destination', 'Travelling Cost', 'Travelling Claim', 'DA', 'DA Claim', 'Total Claim']
                ];
                
                expenses.forEach(expense => {
                    excelData.push([
                        expense.date,
                        expense.destination,
                        expense.travellingCostSum,
                        expense.travellingClaimSum,
                        expense.daSum,
                        expense.daClaimSum,
                        expense.totalClaim
                    ]);
                });
                
                // Add totals row
                const totalTravellingCost = expenses.reduce((sum, exp) => sum + exp.travellingCostSum, 0);
                const totalTravellingClaim = expenses.reduce((sum, exp) => sum + exp.travellingClaimSum, 0);
                const totalDA = expenses.reduce((sum, exp) => sum + exp.daSum, 0);
                const totalDAClaim = expenses.reduce((sum, exp) => sum + exp.daClaimSum, 0);
                const totalClaim = expenses.reduce((sum, exp) => sum + exp.totalClaim, 0);
                
                excelData.push([
                    'TOTAL',
                    '',
                    totalTravellingCost,
                    totalTravellingClaim,
                    totalDA,
                    totalDAClaim,
                    totalClaim
                ]);
                
                // Create worksheet
                const ws = XLSX.utils.aoa_to_sheet(excelData);
                
                // Create workbook
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Expenses');
                
                // Generate file and download
                const fileName = `${currentFile.name || 'expenses'}.xlsx`;
                XLSX.writeFile(wb, fileName);
            });

            // Save file button
            saveFileBtn.addEventListener('click', function() {
                if (expenses.length === 0) {
                    alert('No data to save');
                    return;
                }
                
                const fileName = prompt('Enter a name for your file:', currentFile.name);
                
                if (fileName) {
                    saveCurrentFileAs(fileName);
                    alert(`File "${fileName}" saved successfully!`);
                }
            });

            // New file button
            newFileBtn.addEventListener('click', function() {
                if (expenses.length > 0 && confirm('Are you sure you want to create a new file? The current file will be saved automatically.')) {
                    saveCurrentFile();
                    createNewFile();
                } else if (expenses.length === 0) {
                    createNewFile();
                }
            });

            // Save current file with a new name
            function saveCurrentFileAs(fileName) {
                if (expenses.length === 0) return;
                
                const fileId = currentFile.id || Date.now();
                const now = new Date();
                
                const fileToSave = {
                    id: fileId,
                    name: fileName,
                    data: [...expenses],
                    createdAt: currentFile.createdAt || now,
                    updatedAt: now,
                    type: 'expense'
                };
                
                // Check if this file already exists
                const existingFileIndex = files.findIndex(file => file.id === fileId);
                if (existingFileIndex !== -1) {
                    files[existingFileIndex] = fileToSave;
                } else {
                    files.push(fileToSave);
                }
                
                // Update current file name
                currentFile.name = fileName;
                
                // Add to history
                addToHistory(fileId);
                
                // Save to localStorage
                localStorage.setItem('expenseFiles', JSON.stringify(files));
                localStorage.setItem('expenseHistory', JSON.stringify(history));
            }

            // Save current file (without changing name)
            function saveCurrentFile() {
                if (!currentFile.id || expenses.length === 0) return;
                
                const now = new Date();
                
                const fileToSave = {
                    ...currentFile,
                    data: [...expenses],
                    updatedAt: now
                };
                
                // Update in files array
                const existingFileIndex = files.findIndex(file => file.id === currentFile.id);
                if (existingFileIndex !== -1) {
                    files[existingFileIndex] = fileToSave;
                } else {
                    files.push(fileToSave);
                }
                
                // Add to history
                addToHistory(currentFile.id);
                
                // Save to localStorage
                localStorage.setItem('expenseFiles', JSON.stringify(files));
                localStorage.setItem('expenseHistory', JSON.stringify(history));
            }

            // Add file to history
            function addToHistory(fileId) {
                // Remove if already in history
                history = history.filter(item => item.id !== fileId);
                
                // Add to beginning of history
                history.unshift({
                    id: fileId,
                    accessedAt: new Date()
                });
                
                // Keep only last 10 items
                if (history.length > 10) {
                    history = history.slice(0, 10);
                }
            }

            // Create a new empty file
            function createNewFile() {
                currentFile = {
                    id: Date.now(),
                    name: 'Untitled',
                    data: [],
                    createdAt: new Date(),
                    type: 'expense'
                };
                
                expenses = [];
                renderExpenseTable();
                updateSummary();
                
                // Clear form
                expenseForm.reset();
                stayChargeGroup.style.display = 'none';
            }

            // Render history list
            function renderHistoryList() {
                historyList.innerHTML = '';
                
                if (history.length === 0) {
                    historyList.innerHTML = '<li>No recent files</li>';
                } else {
                    history.forEach(item => {
                        const file = files.find(f => f.id === item.id);
                        if (file) {
                            const li = document.createElement('li');
                            li.style.padding = '10px';
                            li.style.borderBottom = '1px solid #eee';
                            li.style.cursor = 'pointer';
                            
                            li.innerHTML = `
                                <strong>${file.name}</strong><br>
                                <small>Last accessed: ${new Date(item.accessedAt).toLocaleString()}</small>
                            `;
                            
                            li.addEventListener('click', function() {
                                loadFile(file.id);
                                loadPage(file.type === 'custom' ? 'custom' : 'home');
                            });
                            
                            historyList.appendChild(li);
                        }
                    });
                }
            }

            // Render files list
            function renderFilesList() {
                filesTableBody.innerHTML = '';
                
                if (files.length === 0) {
                    noFilesMessage.style.display = 'block';
                    filesTable.style.display = 'none';
                } else {
                    noFilesMessage.style.display = 'none';
                    filesTable.style.display = 'table';
                    
                    files.forEach(file => {
                        const row = document.createElement('tr');
                        
                        row.innerHTML = `
                            <td>${file.name}</td>
                            <td>${new Date(file.createdAt).toLocaleDateString()}</td>
                            <td>
                                <div class="action-buttons">
                                    <button class="action-btn rename-btn" data-id="${file.id}">Rename</button>
                                    <button class="action-btn delete-btn" data-id="${file.id}">Delete</button>
                                    <button class="action-btn edit-btn" data-id="${file.id}">Edit</button>
                                </div>
                            </td>
                        `;
                        
                        filesTableBody.appendChild(row);
                    });
                    
                    // Add event listeners to action buttons
                    document.querySelectorAll('.rename-btn').forEach(btn => {
                        btn.addEventListener('click', function() {
                            const id = parseInt(this.getAttribute('data-id'));
                            renameFile(id);
                        });
                    });
                    
                    document.querySelectorAll('.delete-btn').forEach(btn => {
                        btn.addEventListener('click', function() {
                            const id = parseInt(this.getAttribute('data-id'));
                            deleteFile(id);
                        });
                    });
                    
                    document.querySelectorAll('.edit-btn').forEach(btn => {
                        btn.addEventListener('click', function() {
                            const id = parseInt(this.getAttribute('data-id'));
                            loadFile(id);
                            loadPage('home');
                        });
                    });
                }
            }

            // Rename file
            function renameFile(id) {
                const file = files.find(f => f.id === id);
                if (!file) return;
                
                const newName = prompt('Enter new name for the file:', file.name);
                
                if (newName && newName.trim() !== '') {
                    file.name = newName.trim();
                    file.updatedAt = new Date();
                    
                    // If this is the current file, update the current file reference
                    if (currentFile.id === file.id) {
                        currentFile.name = newName.trim();
                    }
                    
                    // Save to localStorage
                    localStorage.setItem('expenseFiles', JSON.stringify(files));
                    
                    // Refresh files list
                    renderFilesList();
                }
            }

            // Delete file
            function deleteFile(id) {
                if (confirm('Are you sure you want to delete this file? This cannot be undone.')) {
                    // Remove from files array
                    files = files.filter(f => f.id !== id);
                    
                    // Remove from history
                    history = history.filter(item => item.id !== id);
                    
                    // If this is the current file, create a new empty file
                    if (currentFile.id === id) {
                        createNewFile();
                    }
                    
                    // Save to localStorage
                    localStorage.setItem('expenseFiles', JSON.stringify(files));
                    localStorage.setItem('expenseHistory', JSON.stringify(history));
                    
                    // Refresh files list
                    renderFilesList();
                }
            }

            // Load file
            function loadFile(id, showAlert = true) {
                const file = files.find(f => f.id === id);
                if (!file) return;
                
                if (file.type === 'expense') {
                    currentFile = file;
                    expenses = file.data;
                    
                    // Render the data
                    renderExpenseTable();
                    updateSummary();
                } else if (file.type === 'custom') {
                    currentCustomFile = file;
                    customColumns = file.columns;
                    customTableData = file.data;
                    
                    // Render the custom table
                    customTableNameInput.value = file.name;
                    customTableTitle.textContent = file.name;
                    renderColumnList();
                    renderCustomTableForm();
                    renderCustomTable();
                    customTableContainer.style.display = 'block';
                }
                
                // Add to history
                addToHistory(file.id);
                localStorage.setItem('expenseHistory', JSON.stringify(history));
                
                // Show success message
                if (showAlert) {
                    alert(`File "${file.name}" loaded successfully!`);
                }
            }

            // Custom Table Builder Functionality
            columnTypeSelect.addEventListener('change', function() {
                const selectedType = this.value;
                if (selectedType === 'radio' || selectedType === 'select' || selectedType === 'checkbox') {
                    optionsGroup.style.display = 'block';
                } else {
                    optionsGroup.style.display = 'none';
                }
            });

            addColumnBtn.addEventListener('click', function() {
                const name = columnNameInput.value.trim();
                const type = columnTypeSelect.value;
                
                if (!name) {
                    alert('Please enter a column name');
                    return;
                }
                
                const column = {
                    id: Date.now(),
                    name,
                    type,
                    options: type === 'radio' || type === 'select' || type === 'checkbox' ? 
                        columnOptionsInput.value.split(',').map(opt => opt.trim()) : []
                };
                
                customColumns.push(column);
                renderColumnList();
                
                // Clear inputs
                columnNameInput.value = '';
                columnOptionsInput.value = '';
                optionsGroup.style.display = 'none';
                columnTypeSelect.value = 'text';
            });

            function renderColumnList() {
                columnList.innerHTML = '';
                
                customColumns.forEach((column, index) => {
                    const columnItem = document.createElement('div');
                    columnItem.className = 'column-item';
                    
                    columnItem.innerHTML = `
                        <span><strong>${column.name}</strong> (${column.type})</span>
                        <div class="column-actions">
                            <button class="action-btn delete-btn" data-index="${index}">Delete</button>
                        </div>
                    `;
                    
                    columnList.appendChild(columnItem);
                });
                
                // Add event listeners to delete buttons
                document.querySelectorAll('.column-item .delete-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const index = parseInt(this.getAttribute('data-index'));
                        customColumns.splice(index, 1);
                        renderColumnList();
                    });
                });
            }

            createCustomTableBtn.addEventListener('click', function() {
                const tableName = customTableNameInput.value.trim();
                
                if (!tableName) {
                    alert('Please enter a table name');
                    return;
                }
                
                if (customColumns.length === 0) {
                    alert('Please add at least one column');
                    return;
                }
                
                currentCustomFile.name = tableName;
                currentCustomFile.columns = [...customColumns];
                currentCustomFile.data = [];
                currentCustomFile.createdAt = new Date();
                currentCustomFile.updatedAt = new Date();
                
                customTableTitle.textContent = tableName;
                
                // Render the form and table
                renderCustomTableForm();
                renderCustomTable();
                
                // Show the table container
                customTableContainer.style.display = 'block';
            });

            function renderCustomTableForm() {
                customTableForm.innerHTML = '';
                
                customColumns.forEach(column => {
                    const formGroup = document.createElement('div');
                    formGroup.className = 'form-group';
                    
                    const label = document.createElement('label');
                    label.textContent = column.name;
                    label.htmlFor = `custom-${column.id}`;
                    
                    let input;
                    
                    switch(column.type) {
                        case 'text':
                        case 'number':
                        case 'date':
                            input = document.createElement('input');
                            input.type = column.type;
                            input.id = `custom-${column.id}`;
                            input.className = 'form-control';
                            break;
                        case 'checkbox':
                            input = document.createElement('div');
                            column.options.forEach(option => {
                                const checkboxDiv = document.createElement('div');
                                const checkbox = document.createElement('input');
                                checkbox.type = 'checkbox';
                                checkbox.id = `custom-${column.id}-${option.replace(/\s+/g, '-')}`;
                                checkbox.name = `custom-${column.id}`;
                                checkbox.value = option;
                                
                                const checkboxLabel = document.createElement('label');
                                checkboxLabel.htmlFor = `custom-${column.id}-${option.replace(/\s+/g, '-')}`;
                                checkboxLabel.textContent = option;
                                
                                checkboxDiv.appendChild(checkbox);
                                checkboxDiv.appendChild(checkboxLabel);
                                input.appendChild(checkboxDiv);
                            });
                            break;
                        case 'radio':
                            input = document.createElement('div');
                            column.options.forEach(option => {
                                const radioDiv = document.createElement('div');
                                const radio = document.createElement('input');
                                radio.type = 'radio';
                                radio.id = `custom-${column.id}-${option.replace(/\s+/g, '-')}`;
                                radio.name = `custom-${column.id}`;
                                radio.value = option;
                                
                                const radioLabel = document.createElement('label');
                                radioLabel.htmlFor = `custom-${column.id}-${option.replace(/\s+/g, '-')}`;
                                radioLabel.textContent = option;
                                
                                radioDiv.appendChild(radio);
                                radioDiv.appendChild(radioLabel);
                                input.appendChild(radioDiv);
                            });
                            break;
                        case 'select':
                            input = document.createElement('select');
                            input.id = `custom-${column.id}`;
                            input.className = 'form-control';
                            column.options.forEach(option => {
                                const optionEl = document.createElement('option');
                                optionEl.value = option;
                                optionEl.textContent = option;
                                input.appendChild(optionEl);
                            });
                            break;
                    }
                    
                    formGroup.appendChild(label);
                    formGroup.appendChild(input);
                    customTableForm.appendChild(formGroup);
                });
                
                const submitBtn = document.createElement('button');
                submitBtn.type = 'button';
                submitBtn.className = 'btn-submit';
                submitBtn.textContent = 'Add Row';
                submitBtn.addEventListener('click', addCustomRow);
                
                customTableForm.appendChild(submitBtn);
            }

            function addCustomRow() {
                const rowData = {};
                
                customColumns.forEach(column => {
                    const input = document.getElementById(`custom-${column.id}`);
                    
                    switch(column.type) {
                        case 'text':
                        case 'number':
                        case 'date':
                        case 'select':
                            rowData[column.id] = input.value;
                            break;
                        case 'checkbox':
                            const checkboxes = document.querySelectorAll(`input[name="custom-${column.id}"]:checked`);
                            rowData[column.id] = Array.from(checkboxes).map(cb => cb.value).join(', ');
                            break;
                        case 'radio':
                            const radio = document.querySelector(`input[name="custom-${column.id}"]:checked`);
                            rowData[column.id] = radio ? radio.value : '';
                            break;
                    }
                    
                    // Clear the input
                    if (column.type !== 'radio' && column.type !== 'checkbox') {
                        input.value = '';
                    } else if (column.type === 'checkbox') {
                        document.querySelectorAll(`input[name="custom-${column.id}"]`).forEach(cb => {
                            cb.checked = false;
                        });
                    }
                });
                
                rowData.id = Date.now();
                customTableData.push(rowData);
                
                // Update current custom file data
                currentCustomFile.data = customTableData;
                currentCustomFile.updatedAt = new Date();
                
                // Save to localStorage
                saveCustomFile();
                
                // Re-render the table
                renderCustomTable();
            }

            function renderCustomTable() {
                customTableHead.innerHTML = '';
                customTableBody.innerHTML = '';
                
                // Create header row
                const headerRow = document.createElement('tr');
                
                customColumns.forEach(column => {
                    const th = document.createElement('th');
                    th.textContent = column.name;
                    headerRow.appendChild(th);
                });
                
                // Add actions header
                const actionsTh = document.createElement('th');
                actionsTh.textContent = 'Actions';
                headerRow.appendChild(actionsTh);
                
                customTableHead.appendChild(headerRow);
                
                // Create data rows
                if (customTableData.length === 0) {
                    const emptyRow = document.createElement('tr');
                    const emptyCell = document.createElement('td');
                    emptyCell.colSpan = customColumns.length + 1;
                    emptyCell.textContent = 'No data added yet';
                    emptyCell.style.textAlign = 'center';
                    emptyRow.appendChild(emptyCell);
                    customTableBody.appendChild(emptyRow);
                } else {
                    customTableData.forEach(row => {
                        const tr = document.createElement('tr');
                        
                        customColumns.forEach(column => {
                            const td = document.createElement('td');
                            td.textContent = row[column.id] || '';
                            tr.appendChild(td);
                        });
                        
                        // Add actions cell
                        const actionsTd = document.createElement('td');
                        actionsTd.innerHTML = `
                            <div class="action-buttons">
                                <button class="action-btn edit-btn" data-id="${row.id}">Edit</button>
                                <button class="action-btn delete-btn" data-id="${row.id}">Delete</button>
                            </div>
                        `;
                        tr.appendChild(actionsTd);
                        
                        customTableBody.appendChild(tr);
                    });
                }
                
                // Add event listeners to action buttons
                document.querySelectorAll('#customTable .edit-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const id = parseInt(this.getAttribute('data-id'));
                        editCustomRow(id);
                    });
                });
                
                document.querySelectorAll('#customTable .delete-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const id = parseInt(this.getAttribute('data-id'));
                        deleteCustomRow(id);
                    });
                });
            }

            function editCustomRow(id) {
                const row = customTableData.find(r => r.id === id);
                if (!row) return;
                
                // Fill the form with the row data
                customColumns.forEach(column => {
                    const input = document.getElementById(`custom-${column.id}`);
                    
                    switch(column.type) {
                        case 'text':
                        case 'number':
                        case 'date':
                        case 'select':
                            if (input) input.value = row[column.id] || '';
                            break;
                        case 'checkbox':
                            const values = row[column.id] ? row[column.id].split(', ') : [];
                            document.querySelectorAll(`input[name="custom-${column.id}"]`).forEach(cb => {
                                cb.checked = values.includes(cb.value);
                            });
                            break;
                        case 'radio':
                            const radio = document.querySelector(`input[name="custom-${column.id}"][value="${row[column.id]}"]`);
                            if (radio) radio.checked = true;
                            break;
                    }
                });
                
                // Remove the row from the array
                customTableData = customTableData.filter(r => r.id !== id);
                
                // Update current custom file data
                currentCustomFile.data = customTableData;
                currentCustomFile.updatedAt = new Date();
                
                // Save to localStorage
                saveCustomFile();
                
                // Re-render the table
                renderCustomTable();
            }

            function deleteCustomRow(id) {
                if (confirm('Are you sure you want to delete this row?')) {
                    customTableData = customTableData.filter(r => r.id !== id);
                    
                    // Update current custom file data
                    currentCustomFile.data = customTableData;
                    currentCustomFile.updatedAt = new Date();
                    
                    // Save to localStorage
                    saveCustomFile();
                    
                    // Re-render the table
                    renderCustomTable();
                }
            }

            function saveCustomFile() {
                if (!currentCustomFile.id) return;
                
                const now = new Date();
                
                const fileToSave = {
                    ...currentCustomFile,
                    data: [...customTableData],
                    columns: [...customColumns],
                    updatedAt: now
                };
                
                // Check if this file already exists
                const existingFileIndex = files.findIndex(file => file.id === currentCustomFile.id);
                if (existingFileIndex !== -1) {
                    files[existingFileIndex] = fileToSave;
                } else {
                    files.push(fileToSave);
                }
                
                // Add to history
                addToHistory(currentCustomFile.id);
                
                // Save to localStorage
                localStorage.setItem('expenseFiles', JSON.stringify(files));
                localStorage.setItem('expenseHistory', JSON.stringify(history));
            }

            downloadCustomExcelBtn.addEventListener('click', function() {
                if (customTableData.length === 0) {
                    alert('No data to export');
                    return;
                }
                
                // Prepare data for Excel
                const excelData = [
                    customColumns.map(col => col.name)
                ];
                
                customTableData.forEach(row => {
                    const rowData = customColumns.map(col => row[col.id] || '');
                    excelData.push(rowData);
                });
                
                // Create worksheet
                const ws = XLSX.utils.aoa_to_sheet(excelData);
                
                // Create workbook
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, currentCustomFile.name);
                
                // Generate file and download
                const fileName = `${currentCustomFile.name || 'custom_data'}.xlsx`;
                XLSX.writeFile(wb, fileName);
            });

            saveCustomFileBtn.addEventListener('click', function() {
                if (customTableData.length === 0) {
                    alert('No data to save');
                    return;
                }
                
                const fileName = prompt('Enter a name for your custom file:', currentCustomFile.name);
                
                if (fileName) {
                    currentCustomFile.name = fileName;
                    saveCustomFile();
                    alert(`File "${fileName}" saved successfully!`);
                }
            });

            // Search functionality
            sidebarSearch.addEventListener('input', function() {
                const searchTerm = this.value.toLowerCase();
                if (searchTerm === '') {
                    renderExpenseTable();
                    return;
                }
                
                const filteredExpenses = expenses.filter(expense => 
                    expense.date.toLowerCase().includes(searchTerm) || 
                    expense.destination.toLowerCase().includes(searchTerm)
                );
                
                renderFilteredExpenses(filteredExpenses);
            });

            function renderFilteredExpenses(filteredExpenses) {
                expenseTableBody.innerHTML = '';
                
                if (filteredExpenses.length === 0) {
                    const row = document.createElement('tr');
                    row.innerHTML = `<td colspan="8" style="text-align: center;">No matching expenses found</td>`;
                    expenseTableBody.appendChild(row);
                    return;
                }
                
                filteredExpenses.forEach(expense => {
                    const row = document.createElement('tr');
                    
                    row.innerHTML = `
                        <td>${expense.date}</td>
                        <td>${formatMultiline(expense.destination)}</td>
                        <td>${formatMultiline(expense.travellingCost)}<br><strong>${expense.travellingCostSum.toFixed(2)}</strong></td>
                        <td>${formatMultiline(expense.travellingClaim)}<br><strong>${expense.travellingClaimSum.toFixed(2)}</strong></td>
                        <td>${formatMultiline(expense.da)}<br><strong>${expense.daSum.toFixed(2)}</strong></td>
                        <td>${formatMultiline(expense.daClaim)}<br><strong>${expense.daClaimSum.toFixed(2)}</strong></td>
                        <td><strong>${expense.totalClaim.toFixed(2)}</strong></td>
                        <td>
                            <div class="action-buttons">
                                <button class="action-btn edit-btn" data-id="${expense.id}">Edit</button>
                                <button class="action-btn delete-btn" data-id="${expense.id}">Delete</button>
                            </div>
                        </td>
                    `;
                    
                    expenseTableBody.appendChild(row);
                });
                
                // Add event listeners to action buttons
                document.querySelectorAll('.edit-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const id = parseInt(this.getAttribute('data-id'));
                        editExpense(id);
                    });
                });
                
                document.querySelectorAll('.delete-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const id = parseInt(this.getAttribute('data-id'));
                        deleteExpense(id);
                    });
                });
            }

            filesSearch.addEventListener('input', function() {
                const searchTerm = this.value.toLowerCase();
                if (searchTerm === '') {
                    renderFilesList();
                    return;
                }
                
                const filteredFiles = files.filter(file => 
                    file.name.toLowerCase().includes(searchTerm)
                );
                
                renderFilteredFiles(filteredFiles);
            });

            function renderFilteredFiles(filteredFiles) {
                filesTableBody.innerHTML = '';
                
                if (filteredFiles.length === 0) {
                    noFilesMessage.style.display = 'block';
                    filesTable.style.display = 'none';
                } else {
                    noFilesMessage.style.display = 'none';
                    filesTable.style.display = 'table';
                    
                    filteredFiles.forEach(file => {
                        const row = document.createElement('tr');
                        
                        row.innerHTML = `
                            <td>${file.name}</td>
                            <td>${new Date(file.createdAt).toLocaleDateString()}</td>
                            <td>
                                <div class="action-buttons">
                                    <button class="action-btn rename-btn" data-id="${file.id}">Rename</button>
                                    <button class="action-btn delete-btn" data-id="${file.id}">Delete</button>
                                    <button class="action-btn edit-btn" data-id="${file.id}">Edit</button>
                                </div>
                            </td>
                        `;
                        
                        filesTableBody.appendChild(row);
                    });
                    
                    // Add event listeners to action buttons
                    document.querySelectorAll('.rename-btn').forEach(btn => {
                        btn.addEventListener('click', function() {
                            const id = parseInt(this.getAttribute('data-id'));
                            renameFile(id);
                        });
                    });
                    
                    document.querySelectorAll('.delete-btn').forEach(btn => {
                        btn.addEventListener('click', function() {
                            const id = parseInt(this.getAttribute('data-id'));
                            deleteFile(id);
                        });
                    });
                    
                    document.querySelectorAll('.edit-btn').forEach(btn => {
                        btn.addEventListener('click', function() {
                            const id = parseInt(this.getAttribute('data-id'));
                            loadFile(id);
                            loadPage('home');
                        });
                    });
                }
            }

            // Initialize the app
            init();
        });

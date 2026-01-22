/**
 * Fitters Dashboard Calendar Widget JavaScript
 * 
 * Handles the functionality for the fitters calendar widget including:
 * - Calendar navigation and date management
 * - Fitter and order filtering
 * - Event display and modal management
 * - Zoho CRM integration for Contacts (Fitters) and Ordering modules
 * - Responsive calendar layout
 */

document.addEventListener('DOMContentLoaded', function() {
    console.log('Fitters Dashboard Calendar Widget: Initializing...');
    
    // Initialize Zoho CRM SDK
    initializeZohoCRM();
    
    // Initialize calendar
    initializeCalendar();
    
    // Initialize event listeners
    initializeEventListeners();
    
    // Initialize tabs
    initializeTabs();
    
    // Initialize edit modal
    initializeEditModal();
    
    // Load initial data
    loadCalendarData();
  });
  
  // =============================================
  // Global Variables
  // =============================================
  
  let currentDate = new Date();
  let weekOffset = 0;
  let fitters = [];
  let orders = [];
  let holidays = []; // Store holiday data
  let filters = {
    fitter: '',
    customer: '',
    jobType: '',
    skillset: '',
    skillsShortage: [],
    postcode: ''
  };
  
  // =============================================
  // Zoho CRM Integration
  // =============================================
  
  async function initializeZohoCRM() {
    try {
      console.log('Initializing Zoho CRM SDK...');
      
      // Wait a bit for the SDK to load
      await new Promise(resolve => setTimeout(resolve, 500));
      
      // Initialize Zoho CRM SDK
      if (typeof ZOHO !== 'undefined') {
        console.log('ZOHO object found, setting up embedded app...');
        
        // Subscribe to PageLoad event before initializing
        ZOHO.embeddedApp.on("PageLoad", function(data) {
          console.log('PageLoad event triggered:', data);
          // Load real data when embedded in Zoho CRM
          loadRealData().catch(error => {
            console.error('Error loading real data from PageLoad:', error);
            // Don't show error in banner during initial load, just keep loading state
            document.getElementById('appointmentCount').textContent = 'Loading...';
            document.getElementById('eventCount').textContent = 'Loading...';
          });
        });
        
        await ZOHO.embeddedApp.init();
        console.log('Zoho CRM SDK initialized successfully');
        
        // Try to load data immediately if SDK is ready
        if (ZOHO.CRM && ZOHO.CRM.API && ZOHO.CRM.API.coql) {
          console.log('Zoho CRM API is ready, attempting to load data...');
          loadRealData().catch(error => {
            console.error('Error loading real data after init:', error);
            // Don't show error in banner during initial load, just keep loading state
            document.getElementById('appointmentCount').textContent = 'Loading...';
            document.getElementById('eventCount').textContent = 'Loading...';
          });
        } else {
          console.log('Zoho CRM API not fully ready, will retry...');
          // Retry after a short delay
          setTimeout(() => {
            if (ZOHO.CRM && ZOHO.CRM.API && ZOHO.CRM.API.coql) {
              loadRealData().catch(error => {
                console.error('Error loading real data on retry:', error);
              });
            }
          }, 1000);
        }
      } else {
        console.log('Zoho CRM SDK not available, will use demo mode');
      }
    } catch (error) {
      console.error('Error initializing Zoho CRM SDK:', error);
      console.log('Continuing with demo data...');
    }
  }
  
  // =============================================
  // Helper Functions
  // =============================================
  
  // Format currency with commas
  function formatCurrency(amount) {
    if (!amount || amount === '£0.00' || amount === 'Not provided') return '£0.00';
    
    // Remove £ symbol and convert to number
    const numericAmount = parseFloat(amount.toString().replace(/[£,]/g, ''));
    if (isNaN(numericAmount)) return '£0.00';
    
    // Format with commas and add £ symbol
    return '£' + numericAmount.toLocaleString('en-GB', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }
  
  // Standardize empty data display
  function getDisplayValue(value, fallback = 'Missing Data') {
    if (!value || value === '' || value === 'N/A' || value === 'Not Set' || value === 'Not specified' || value === 'Not provided') {
      return fallback;
    }
    return value;
  }

  // Check if data is available for styling
  function isDataAvailable(value) {
    return value && value !== '' && value !== 'N/A' && value !== 'Not Set' && value !== 'Not specified' && value !== 'Not provided';
  }

  // Get CSS class for data status
  function getDataStatusClass(value) {
    return isDataAvailable(value) ? 'text-data-available' : 'text-data-missing';
  }
  
  // Generate Zoho CRM order URL
  function getOrderUrl(orderId) {
    try {
      // Use the correct Zoho CRM URL format for Orders module
      const orgId = ZOHO.CRM.API.getOrgId ? ZOHO.CRM.API.getOrgId() : '2579410';
      return `https://crm.zoho.com/crm/org${orgId}/tab/CustomModule13/${orderId}/canvas/167246000058981119`;
    } catch (error) {
      console.error('Error generating order URL:', error);
      // Fallback URL with default orgId
      return `https://crm.zoho.com/crm/org2579410/tab/CustomModule13/${orderId}/canvas/167246000058981119`;
    }
  }
  
  // Generate Zoho CRM contact URL
  function getContactUrl(contactId) {
    try {
      // Use the correct Zoho CRM URL format for Contacts module
      const orgId = ZOHO.CRM.API.getOrgId ? ZOHO.CRM.API.getOrgId() : '2579410';
      return `https://crm.zoho.com/crm/org${orgId}/tab/Contacts/${contactId}/canvas/167246000044297004`;
    } catch (error) {
      console.error('Error generating contact URL:', error);
      // Fallback URL with default orgId
      return `https://crm.zoho.com/crm/org2579410/tab/Contacts/${contactId}/canvas/167246000044297004`;
    }
  }
  
  function selectFitter(fitterName) {
    const fitterSearch = document.getElementById('fitterSearch');
    const fitterFilter = document.getElementById('fitterFilter');
    const fitterDropdown = document.getElementById('fitterDropdown');
    
    fitterSearch.value = fitterName;
    fitterFilter.value = fitterName;
    fitterDropdown.classList.remove('show');
    
    // Update selected state
    const options = fitterDropdown.querySelectorAll('.select-option');
    options.forEach(option => {
      option.classList.remove('selected');
      if (option.textContent === fitterName) {
        option.classList.add('selected');
      }
    });
    
    handleFilterChange();
  }
  
  function selectCustomer(customerName) {
    const customerSearch = document.getElementById('customerSearch');
    const customerFilter = document.getElementById('customerFilter');
    const customerDropdown = document.getElementById('customerDropdown');
    
    customerSearch.value = customerName;
    customerFilter.value = customerName;
    customerDropdown.classList.remove('show');
    
    // Update selected state
    const options = customerDropdown.querySelectorAll('.select-option');
    options.forEach(option => {
      option.classList.remove('selected');
      if (option.textContent === customerName) {
        option.classList.add('selected');
      }
    });
    
    handleFilterChange();
  }
  
  function parseSkillsShortageString(str) {
    // Split by comma but only if not inside parentheses
    const result = [];
    let current = '';
    let parenCount = 0;
    
    for (let i = 0; i < str.length; i++) {
      const char = str[i];
      if (char === '(') parenCount++;
      else if (char === ')') parenCount--;
      else if (char === ',' && parenCount === 0) {
        if (current.trim()) {
          result.push(current.trim());
          current = '';
        }
        continue;
      }
      current += char;
    }
    
    if (current.trim()) {
      result.push(current.trim());
    }
    
    return result.filter(item => item.length > 0);
  }
  
  // =============================================
  // Skills Shortage Multi-Select
  // =============================================
  
   function initializeCustomMultiSelect() {
     const container = document.querySelector('.custom-multi-select');
     const trigger = document.getElementById('skillsShortageTrigger');
     const dropdown = document.getElementById('skillsShortageDropdown');
     
     if (!container || !trigger || !dropdown) {
       return;
     }
     
     const text = trigger.querySelector('.multi-select-text');
     if (!text) {
       return;
     }
    
     // Toggle dropdown
     trigger.addEventListener('click', (e) => {
       e.preventDefault();
       e.stopPropagation();
       container.classList.toggle('open');
     });
     
     // Fallback click handler for the entire container
     container.addEventListener('click', (e) => {
       if (e.target === trigger || trigger.contains(e.target)) {
         e.preventDefault();
         e.stopPropagation();
         container.classList.toggle('open');
       }
     });
    
    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
      if (!container.contains(e.target)) {
        container.classList.remove('open');
      }
    });
    
    // Handle checkbox changes
     dropdown.addEventListener('change', (e) => {
       if (e.target.type === 'checkbox' && e.target.classList.contains('skills-shortage-checkbox')) {
         updateMultiSelectText();
         handleFilterChange();
       }
     });
     
     // Also handle clicks on the option divs
     dropdown.addEventListener('click', (e) => {
       if (e.target.classList.contains('multi-select-option') || e.target.tagName === 'LABEL') {
         const checkbox = e.target.querySelector('input[type="checkbox"]') || e.target.previousElementSibling;
         if (checkbox && checkbox.classList.contains('skills-shortage-checkbox')) {
           checkbox.checked = !checkbox.checked;
           updateMultiSelectText();
           handleFilterChange();
         }
       }
     });
    
    function updateMultiSelectText() {
      const checkedBoxes = dropdown.querySelectorAll('input[type="checkbox"]:checked');
      if (checkedBoxes.length === 0) {
        text.textContent = 'No Exclusions';
      } else if (checkedBoxes.length === 1) {
        text.textContent = checkedBoxes[0].value;
      } else {
        text.textContent = `${checkedBoxes.length} selected`;
      }
    }
  }
  
  // =============================================
  // Calendar Initialization
  // =============================================
  
  function initializeCalendar() {
    // Set current date
    currentDate = new Date();
    weekOffset = 0;
    
    // Initialize filters
    filters = {
      fitter: '',
      customer: '',
      jobType: '',
      skillset: '',
      skillsShortage: [],
      postcode: ''
    };
    
    // Update today button state
    updateTodayButtonState();
    
    console.log('Calendar initialized with date:', currentDate.toDateString());
  }
  
  // =============================================
  // Event Listeners
  // =============================================
  
  function initializeEventListeners() {
    // Navigation buttons
    document.getElementById('prevWeek').addEventListener('click', goToPreviousWeek);
    document.getElementById('nextWeek').addEventListener('click', goToNextWeek);
    document.getElementById('todayBtn').addEventListener('click', goToToday);
    
    // Filter controls - Add event listeners after DOM is ready
    setTimeout(() => {
      const fitterFilter = document.getElementById('fitterFilter');
      const customerFilter = document.getElementById('customerFilter');
      const jobTypeFilter = document.getElementById('jobTypeFilter');
      const skillsetFilter = document.getElementById('skillsetFilter');
      const skillsShortageFilter = document.getElementById('skillsShortageFilter');
      const postcodeFilter = document.getElementById('postcodeFilter');
      const clearFilters = document.getElementById('clearFilters');
      
      if (fitterFilter) {
        fitterFilter.addEventListener('change', handleFilterChange);
      }
      
      if (customerFilter) {
        customerFilter.addEventListener('change', handleFilterChange);
      }
      
      if (jobTypeFilter) {
        jobTypeFilter.addEventListener('change', handleFilterChange);
      }
      
      if (skillsetFilter) {
        skillsetFilter.addEventListener('change', handleFilterChange);
      }
      
      if (skillsShortageFilter) {
        skillsShortageFilter.addEventListener('change', handleFilterChange);
      }
      
      if (postcodeFilter) {
        postcodeFilter.addEventListener('input', handleFilterChange);
      }
      
      if (clearFilters) {
        clearFilters.addEventListener('click', clearAllFilters);
      }
    }, 500);
    
    // Modal close - Add event listener after DOM is ready
    setTimeout(() => {
      const closeModalBtn = document.getElementById('closeModal');
      const closeModalFooterBtn = document.getElementById('closeModalBtn');
      
      if (closeModalBtn) {
        closeModalBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          closeEventModal();
        });
      }
      
      if (closeModalFooterBtn) {
        closeModalFooterBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          closeEventModal();
        });
      }
    }, 600);
    
    // Close modal when clicking outside - Add event listener after DOM is ready
    setTimeout(() => {
      const eventModal = document.getElementById('eventModal');
      if (eventModal) {
        eventModal.addEventListener('click', function(e) {
      if (e.target === this) {
        closeEventModal();
          }
        });
      }
    }, 700);
    
    // Close modal with ESC key
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        const modal = document.getElementById('eventModal');
        if (modal && modal.style.display === 'block') {
          closeEventModal();
        }
        const createHolidayModal = document.getElementById('createHolidayModal');
        if (createHolidayModal && createHolidayModal.style.display === 'block') {
          closeCreateHolidayModal();
        }
        const sendEmailModal = document.getElementById('sendEmailModal');
        if (sendEmailModal && sendEmailModal.style.display === 'block') {
          closeSendEmailModal();
        }
      }
    });
    
    // Holiday creation modal event listeners
    setTimeout(() => {
      const closeCreateHolidayModalBtn = document.getElementById('closeCreateHolidayModal');
      const cancelCreateHolidayBtn = document.getElementById('cancelCreateHoliday');
      const createHolidayForm = document.getElementById('createHolidayForm');
      
      if (closeCreateHolidayModalBtn) {
        closeCreateHolidayModalBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          closeCreateHolidayModal();
        });
      }
      
      if (cancelCreateHolidayBtn) {
        cancelCreateHolidayBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          closeCreateHolidayModal();
        });
      }
      
      if (createHolidayForm) {
        createHolidayForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          e.stopPropagation();
          await handleCreateHolidaySubmit();
        });
      }
      
      // Close modal when clicking outside
      const createHolidayModal = document.getElementById('createHolidayModal');
      if (createHolidayModal) {
        createHolidayModal.addEventListener('click', function(e) {
          if (e.target === this) {
            closeCreateHolidayModal();
          }
        });
      }
    }, 800);
    
    // Email modal event listeners - only set up once
    setTimeout(() => {
      const closeSendEmailModalBtn = document.getElementById('closeSendEmailModal');
      const cancelSendEmailBtn = document.getElementById('cancelSendEmail');
      const sendEmailForm = document.getElementById('sendEmailForm');
      const emailMessage = document.getElementById('emailMessage');
      const sendEmailModal = document.getElementById('sendEmailModal');
      
      // Only set up listeners if they haven't been set up yet
      if (closeSendEmailModalBtn && !closeSendEmailModalBtn.hasAttribute('data-listener-attached')) {
        closeSendEmailModalBtn.setAttribute('data-listener-attached', 'true');
        closeSendEmailModalBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          closeSendEmailModal();
        });
      }
      
      if (cancelSendEmailBtn && !cancelSendEmailBtn.hasAttribute('data-listener-attached')) {
        cancelSendEmailBtn.setAttribute('data-listener-attached', 'true');
        cancelSendEmailBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          closeSendEmailModal();
        });
      }
      
      if (sendEmailForm && !sendEmailForm.hasAttribute('data-listener-attached')) {
        sendEmailForm.setAttribute('data-listener-attached', 'true');
        sendEmailForm.addEventListener('submit', sendEmailToFitters);
      }
      
      if (emailMessage && !emailMessage.hasAttribute('data-listener-attached')) {
        emailMessage.setAttribute('data-listener-attached', 'true');
        emailMessage.addEventListener('input', updateEmailCharCount);
      }
      
      // Close modal when clicking outside
      if (sendEmailModal && !sendEmailModal.hasAttribute('data-listener-attached')) {
        sendEmailModal.setAttribute('data-listener-attached', 'true');
        sendEmailModal.addEventListener('click', function(e) {
          if (e.target === this) {
            closeSendEmailModal();
          }
        });
      }
    }, 900);
  }
  
  // =============================================
  // Calendar Navigation
  // =============================================
  
  function goToPreviousWeek() {
  weekOffset -= 98; // Move back 14 weeks
  updateCalendar();
  updateTodayButtonState();
  // Load events for new date range
  loadEventsForCurrentDateRange().catch(error => {
    console.error('Error loading events for previous week:', error);
  });
}

function goToNextWeek() {
  weekOffset += 98; // Move forward 14 weeks
  updateCalendar();
  updateTodayButtonState();
  // Load events for new date range
  loadEventsForCurrentDateRange().catch(error => {
    console.error('Error loading events for next week:', error);
  });
}
  
  function goToToday() {
    weekOffset = 0;
    currentDate = new Date();
    updateCalendar();
    updateTodayButtonState();
    // Load events for new date range
    loadEventsForCurrentDateRange().catch(error => {
      console.error('Error loading events for today:', error);
    });
  }

  // Update the Today button state based on current week
  function updateTodayButtonState() {
    const todayBtn = document.getElementById('todayBtn');
    if (!todayBtn) return;
    
    const today = new Date();
    const weeks = getThreeWeekDays();
    const isCurrentWeek = weeks.some(week => week.isToday);
    
    if (isCurrentWeek) {
      todayBtn.classList.add('active');
    } else {
      todayBtn.classList.remove('active');
    }
  }
  
  // =============================================
  // Date Range Formatting
  // =============================================
  
  function formatDateRange(startDate, endDate) {
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    const startDay = start.getDate();
    const startMonth = start.toLocaleDateString('en-US', { month: 'long' });
    const startDayName = start.toLocaleDateString('en-US', { weekday: 'short' });
    
    const endDay = end.getDate();
    const endMonth = end.toLocaleDateString('en-US', { month: 'long' });
    const endDayName = end.toLocaleDateString('en-US', { weekday: 'short' });
    
    // If same month, show: Mon 15th - Sun 5th July
    if (startMonth === endMonth) {
      return `${startDayName} ${startDay} - ${endDayName} ${endDay} ${startMonth}`;
    } else {
      // Different months: Mon 15th July - Sun 5th August
      return `${startDayName} ${startDay} ${startMonth} - ${endDayName} ${endDay} ${endMonth}`;
    }
  }
  
  // =============================================
  // Date Management
  // =============================================
  
  function getThreeWeekDays() {
  const weeks = [];
  const today = new Date();
  
  // Start from the beginning of current week + offset
  const start = new Date(currentDate);
  start.setDate(start.getDate() + weekOffset);
  
  // Get to the Monday of the week
  const dayOfWeek = start.getDay();
  const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // Sunday = 0, Monday = 1
  start.setDate(start.getDate() + mondayOffset);
  
  // Generate 14 weeks starting from last Monday (so this week is in position 2)
  for (let i = -1; i < 13; i++) {
    const monday = new Date(start);
    monday.setDate(start.getDate() + (i * 7));
    
    // Check if this week contains today
    const weekStart = new Date(monday);
    const weekEnd = new Date(monday);
    weekEnd.setDate(monday.getDate() + 6);
    
    const isCurrentWeek = today >= weekStart && today <= weekEnd;
    const dateString = monday.toISOString().split('T')[0]; // YYYY-MM-DD format for Monday
    
    // Get week number
    const weekNumber = getWeekNumber(monday);
    
    // Format date as DD/MM/YY
    const formattedDate = monday.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: '2-digit',
      year: '2-digit'
    });
    
    const label = `WK${weekNumber}\n${formattedDate}`;
    
    weeks.push({
      date: dateString, // Monday's date as the week identifier
      label: label,
      isToday: isCurrentWeek,
      weekStart: weekStart.toISOString().split('T')[0],
      weekEnd: weekEnd.toISOString().split('T')[0],
      fullDate: monday
    });
  }
  
  return weeks;
}

// Helper function to get ISO week number
function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}
  
  function formatDate(date) {
    return date.toISOString().split('T')[0];
  }
  
  function isSameDay(date1, date2) {
    return date1.toDateString() === date2.toDateString();
  }
  
  function getDateRange() {
  const weeks = getThreeWeekDays();
  const start = weeks[0];
  const end = weeks[weeks.length - 1];
  
  return {
    start: start.label.split('\n')[1], // Get the date part (DD/MM/YY)
    end: end.label.split('\n')[1], // Get the date part (DD/MM/YY)
    startDate: start.date,
    endDate: end.date
  };
}
  
  // =============================================
  // Data Management
  // =============================================
  
  function loadCalendarData() {
    showLoading(true);
    
    // Update banner to show loading state
    document.getElementById('eventCount').textContent = 'Loading...';
    
    console.log('loadCalendarData called');
    
    // Check if Zoho SDK is fully ready
    if (typeof ZOHO !== 'undefined' && ZOHO.embeddedApp && ZOHO.CRM && ZOHO.CRM.API && ZOHO.CRM.API.coql) {
      console.log('Zoho CRM API is ready, loading real data...');
    loadRealData().catch(error => {
      console.error('Failed to load real data:', error);
      // Don't show any errors during initial load - just keep loading state
      document.getElementById('eventCount').textContent = 'Loading...';
      showLoading(false);
    });
    } else {
      console.log('Zoho SDK not fully ready, using demo data');
      loadDemoData();
    }
  }
  
  function loadDemoData() {
    console.log('Loading demo data...');
    
     // Demo fitters with special ordering
     fitters = [
       { id: 'unassigned', name: 'Unassigned Fitter', skillset: 'General', skillsShortage: '', postcodesCovered: 'All', workConsistency: 'N/A' },
       { id: 'supply', name: 'Supply Only', skillset: 'Supply', skillsShortage: '', postcodesCovered: 'All', workConsistency: 'N/A' },
       { id: '1', name: 'Aaron Cartwright', skillset: 'Tiled Bathrooms, Wetwall Bathrooms', skillsShortage: 'Complex (Pumps, Wetrooms)', postcodesCovered: 'G1-80', workConsistency: 'High' },
       { id: '2', name: 'David Hamilton', skillset: 'Tiled Bathrooms', skillsShortage: '', postcodesCovered: 'EH1-55', workConsistency: 'Medium' },
       { id: '3', name: 'Stephen Clark', skillset: 'Laminate Kitchens, Mistral Kitchens', skillsShortage: 'Painting (Painting, Wallpapering)', postcodesCovered: 'FK1-20', workConsistency: 'High' },
       { id: '4', name: 'David Shaw', skillset: 'Tiled Bathrooms', skillsShortage: '', postcodesCovered: 'G12, EH6', workConsistency: 'Low' },
       { id: '5', name: 'David McLellan', skillset: 'Tiled Bathrooms, Wetwall Bathrooms', skillsShortage: '', postcodesCovered: 'FK15', workConsistency: 'Medium' },
       { id: '6', name: 'Craig Feeney', skillset: 'Tiled Bathrooms, Wetwall Bathrooms', skillsShortage: '', postcodesCovered: 'ML1-10', workConsistency: 'High' },
       { id: '7', name: 'Kieron Scullion', skillset: 'Laminate Kitchens, Mistral', skillsShortage: '', postcodesCovered: 'G1-80', workConsistency: 'Medium' },
       { id: '8', name: 'Stevie Watson EH', skillset: 'Tiled Bathrooms, Wetwall Bathrooms', skillsShortage: '', postcodesCovered: 'EH1-55', workConsistency: 'Low' }
     ];
    
    // Demo orders
    const today = new Date();
    const currentWeekStart = new Date(today);
    currentWeekStart.setDate(today.getDate() - today.getDay() + 1); // Monday of current week
    
    orders = [
      {
        id: '1',
        title: 'Henderson - Kitchen Installation',
        fitterId: '1',
        fitterName: 'Aaron Cartwright',
        customerName: 'Henderson',
        jobType: 'Kitchen',
        postcode: 'EH10 5PF',
        date: new Date(currentWeekStart.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0], // Next week
        startTime: '09:00',
        endTime: '17:00',
        meetingType: 'Installation',
        meetingStatus: 'Scheduled'
      },
      {
        id: '2',
        title: 'Crawford - Bathroom Installation',
        fitterId: '2',
        fitterName: 'David Hamilton',
        customerName: 'Crawford',
        jobType: 'Bathroom',
        postcode: 'ML6 7QN',
        date: new Date(currentWeekStart.getTime() + 14 * 24 * 60 * 60 * 1000).toISOString().split('T')[0], // Week after next
        startTime: '09:00',
        endTime: '17:00',
        meetingType: 'Installation',
        meetingStatus: 'Scheduled'
      },
      {
        id: '3',
        title: 'Smith - Kitchen Installation',
        fitterId: '3',
        fitterName: 'Stephen Clark',
        customerName: 'Smith',
        jobType: 'Kitchen',
        postcode: 'G12 8AB',
        date: new Date(currentWeekStart.getTime() + 21 * 24 * 60 * 60 * 1000).toISOString().split('T')[0], // 3 weeks from now
        startTime: '09:00',
        endTime: '17:00',
        meetingType: 'Installation',
        meetingStatus: 'Scheduled'
      }
    ];
    
    // Populate filter options
    populateFilterOptions();
    
    // Update calendar display
    updateCalendar();
    
    // Update counts
    updateCounts();
    
    showLoading(false);
    console.log('Demo data loaded successfully');
  }
  
  // =============================================
  // Real Zoho CRM Data Loading
  // =============================================
  
  async function loadRealData() {
    try {
      console.log('Loading real data from Zoho CRM...');
      
      // Check if Zoho SDK is available
      if (typeof ZOHO === 'undefined') {
        console.log('Zoho SDK not available yet, will retry later...');
        throw new Error('Zoho SDK not available');
      }
      
      if (!ZOHO.CRM || !ZOHO.CRM.API) {
        console.log('Zoho CRM API not available yet, will retry later...');
        throw new Error('Zoho CRM API not available');
      }
      
      console.log('Zoho SDK and CRM API are available, proceeding with real data loading...');
      
      // Update loading message
      updateLoadingMessage('Loading fitters...');
      
      // Load fitters from Contacts module
      const fittersResponse = await loadFitters();
      console.log(`Fitters loaded: ${fittersResponse.info.total_count} fitters`);
      
      // Process fitters data
      await processFittersData(fittersResponse);
      
      // Update loading message for orders
      updateLoadingMessage('Loading orders...');
      
      // Load orders for current date range
      await loadEventsForCurrentDateRange();
      
      // Load holidays data
      await loadHolidaysData();
      
      // Populate filter options after all data is processed
      populateFilterOptions();
      
      // Update calendar to show holidays
      updateCalendar();
      
      showLoading(false);
      console.log('Real data loaded successfully');
      
    } catch (error) {
      console.error('Error loading real data:', error);
      throw error; // Re-throw to be caught by the calling function
    }
  }
  
  async function loadEventsForCurrentDateRange() {
  try {
    // Get current week range from calendar
    const weeks = getThreeWeekDays();
    const startDate = weeks[0].weekStart;
    const endDate = weeks[weeks.length - 1].weekEnd;
    
    console.log(`Loading orders for current week range: ${startDate} to ${endDate}`);
    
    // Show loading indicator
    showLoading(true);
    updateLoadingMessage(`Loading orders for weeks ${startDate} to ${endDate}...`);
    
    // Load orders for current week range
    const ordersResponse = await loadOrdersForDateRange(startDate, endDate);
      console.log(`Orders loaded: ${ordersResponse.info.total_count} orders for date range`);
      
      // Process orders data
      processOrdersData(ordersResponse);
      
      // Check if we have any orders
      if (orders.length === 0) {
        showNoEventsMessage();
      } else {
        // Update calendar display
        updateCalendar();
      }
      
      // Hide loading indicator
      showLoading(false);
      
    } catch (error) {
      console.error('Error loading orders for current date range:', error);
      showLoading(false);
      throw error;
    }
  }
  
  function updateLoadingMessage(message) {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) {
      const messageElement = loadingOverlay.querySelector('p');
      if (messageElement) {
        messageElement.textContent = message;
      }
    }
    
    // Also update the calendar navigation to show loading state
    document.getElementById('eventCount').textContent = message;
  }
  
  async function loadFitters() {
  try {
    let allFitters = [];
    let currentPage = 1;
    let hasMoreRecords = true;
    
    console.log('Loading fitters from Contacts module using COQL with criteria...');
    
    // Get fitters with criteria: Fitter_Status equals "Live + Current" or "Live + Current (To Be Audited)" using COQL
    while (hasMoreRecords) {
      console.log(`Fetching Fitters page ${currentPage}...`);
      
      const coqlQuery = {
        "select_query": `select id,First_Name,Last_Name,Email,Home_Phone,Mobile,Fitter_Skillset_2,Skills_They_Can_t_Do,Postcode_Area_New,Work_Consistency,Insurance_Status,PL_Date,PL_Insurance_Number,Fitter_Status,Terms_Masdouk from Contacts where (Fitter_Status = 'Live + Current' or Fitter_Status = 'Live + Current (To Be Audited)') limit 200 offset ${(currentPage - 1) * 200}`
      };
      
      console.log(`COQL Query page ${currentPage}:`, coqlQuery);
      
      const response = await ZOHO.CRM.API.coql(coqlQuery);
      
      console.log(`Page ${currentPage} response:`, response);
      
      if (response.data && response.data.length > 0) {
        allFitters = allFitters.concat(response.data);
        console.log(`Loaded page ${currentPage}: ${response.data.length} fitters`);
      } else {
        console.log(`Page ${currentPage} returned no data`);
      }
      
      // Check if there are more records
      hasMoreRecords = response.info?.more_records || false;
      console.log(`More records available: ${hasMoreRecords}`);
      
      currentPage++;
      
      // Safety check to prevent infinite loops
      if (currentPage > 50) {
        console.warn('Reached maximum page limit for fitters');
        break;
      }
      
      // Small delay between API calls to respect rate limits
      if (hasMoreRecords) {
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }
    
    console.log(`Total fitters loaded: ${allFitters.length}`);
    
    return {
      data: allFitters,
      info: {
        total_count: allFitters.length,
        pages_loaded: currentPage - 1
      }
    };
  } catch (error) {
    console.error('Error loading fitters:', error);
    throw new Error(`Failed to load fitters: ${error.message}`);
  }
}
  
  async function loadOrdersForDateRange(startDate, endDate) {
  try {
    console.log(`Loading orders for date range: ${startDate} to ${endDate}`);
    
    let allOrders = [];
    let currentPage = 1;
    let hasMoreRecords = true;
    
    while (hasMoreRecords) {
      console.log(`Fetching Orders page ${currentPage}...`);
      
      // Use COQL to filter orders within the specified date range using Start_Date
      const coqlQuery = {
        "select_query": `select id,Fitter,Lead,Job_Type,Start_Date,Customer_Name,Postcode,Street,Mobile,Showroom,Designer,Surveyor,Grand_Total,Balance from Ordering where Start_Date >= '${startDate}' and Start_Date <= '${endDate}' limit 200 offset ${(currentPage - 1) * 200}`
      };
      
      console.log(`COQL Query page ${currentPage}:`, coqlQuery);
      
      const response = await ZOHO.CRM.API.coql(coqlQuery);
      
      console.log(`Page ${currentPage} response:`, response);
      
      if (response.data && response.data.length > 0) {
        allOrders = allOrders.concat(response.data);
        console.log(`Loaded page ${currentPage}: ${response.data.length} orders`);
      } else {
        console.log(`Page ${currentPage} returned no data`);
      }
      
      // Check if there are more records
      hasMoreRecords = response.info?.more_records || false;
      console.log(`More records available: ${hasMoreRecords}`);
      
      currentPage++;
      
      // Safety check to prevent infinite loops
      if (currentPage > 50) {
        console.warn('Reached maximum page limit for orders');
        break;
      }
      
      // Small delay between API calls to respect rate limits
      if (hasMoreRecords) {
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }
    
    console.log(`Total orders loaded: ${allOrders.length}`);
    
    return {
      data: allOrders,
      info: {
        total_count: allOrders.length,
        date_range: { start: startDate, end: endDate },
        pages_loaded: currentPage - 1
      }
    };
  } catch (error) {
    console.error('Error loading orders for date range:', error);
    throw new Error(`Failed to load orders for date range: ${error.message}`);
  }
}
  
  async function processFittersData(fittersResponse) {
    console.log('Processing fitters data...');
    
    // Process fitters from Contacts module
    fitters = [];
    
    for (const fitter of fittersResponse.data || []) {
      console.log(`Processing fitter: ${fitter.First_Name} ${fitter.Last_Name}`);
      
      // Handle skillset data - it can be an array or string
      let skillsetData = fitter.Fitter_Skillset_2 || '';
      if (Array.isArray(skillsetData)) {
        skillsetData = skillsetData.join(',');
      }
      
      const fitterData = {
        id: fitter.id,
        firstName: fitter.First_Name || '',
        lastName: fitter.Last_Name || '',
        name: `${fitter.First_Name || ''} ${fitter.Last_Name || ''}`.trim(),
        email: fitter.Email || '',
        phone: fitter.Home_Phone || fitter.Mobile || '',
        skillset: skillsetData,
        skillsShortage: fitter.Skills_They_Can_t_Do || '',
        postcodesCovered: fitter.Postcode_Area_New || '',
        workConsistency: fitter.Work_Consistency || '',
        insuranceStatus: fitter.Insurance_Status || '',
        plDate: fitter.PL_Date || '',
        plInsuranceNo: fitter.PL_Insurance_Number || '',
        termsMasdouk: fitter.Terms_Masdouk || '',
        fitterStatus: fitter.Fitter_Status || ''
      };
      
      fitters.push(fitterData);
      console.log(`Fitter ${fitterData.name} processed with skillset: ${fitterData.skillset}`);
    }
    
    console.log(`Processed ${fitters.length} fitters`);
    console.log(`Sample fitter:`, fitters[0]);
    
    // Store fitters in window object for access by fitters view
    window.fitters = fitters;
    console.log('Stored fitters in window.fitters:', window.fitters.length);
  }
  
  function processOrdersData(ordersResponse) {
  console.log('Processing orders data...');
  
  // Process orders from Ordering module
  orders = (ordersResponse.data || []).map(order => {
    // Find the fitter for this order
    const fitterId = order.Fitter?.id || order.Fitter;
    const fitter = fitters.find(f => f.id === fitterId);
    
    // Get customer name - try Customer_Name field first, then Lead field
    let customerName = 'Unknown Customer';
    
    // Try Customer_Name field first (as mentioned by user)
    if (order.Customer_Name) {
      customerName = order.Customer_Name;
    } else if (order.Lead) {
      if (typeof order.Lead === 'string') {
        customerName = order.Lead;
      } else if (order.Lead.name) {
        customerName = order.Lead.name;
      } else if (order.Lead.id) {
        // If only ID is available, use a generic name
        customerName = `Customer ${order.Lead.id}`;
      }
    }
    
    console.log(`Processing order: ${order.id}, Fitter ID: ${fitterId}, Customer: ${customerName}`);
    
    // Use Start_Date from the order
    let orderDate = order.Start_Date;
    
    // If no Start_Date found, use today as fallback
    if (!orderDate) {
      orderDate = new Date().toISOString().split('T')[0];
    } else {
      // Convert date to YYYY-MM-DD format if it's in a different format
      if (orderDate.includes('T')) {
        orderDate = orderDate.split('T')[0];
      }
    }
    
    const processedPostcode = order.Postcode || '';
    
    // Debug: Log the postcode being processed
    console.log(`Processing order ${order.id}: raw Postcode="${order.Postcode}", processed postcode="${processedPostcode}"`);
    
    return {
      id: order.id,
      title: `${customerName} - ${order.Job_Type || 'Installation'}`,
      fitterId: fitterId,
      fitterName: fitter?.name || `Fitter ${fitterId}`,
      customerName: customerName,
      jobType: order.Job_Type || 'Installation',
      postcode: processedPostcode,
      address: order.Street || '',
      mobile: order.Mobile || '',
      showroom: order.Showroom || '',
      designer: order.Designer?.id || order.Designer || null,
      surveyor: order.Surveyor?.id || order.Surveyor || null,
      goods: order.Grand_Total ? `£${parseFloat(order.Grand_Total).toFixed(2)}` : '£0.00',
      install: order.Balance ? `£${parseFloat(order.Balance).toFixed(2)}` : '£0.00',
      fitterPostcode: fitter?.postcodesCovered || '',
      fitterWorkConsistency: fitter?.workConsistency || '',
      date: orderDate,
      startTime: '09:00',
      endTime: '17:00',
      meetingType: 'Installation',
      meetingStatus: 'Scheduled'
    };
  });
  
  console.log(`Processed ${orders.length} orders for current date range`);
  console.log(`Sample processed order:`, orders[0]);
}
  
  // =============================================
  // Postcode Coverage Logic
  // =============================================
  
  function matchesPostcodeRange(inputPostcode, coverageAreas) {
    if (!inputPostcode || !coverageAreas) {
      return false;
    }
    
    // Ensure coverageAreas is a string
    const coverageString = String(coverageAreas);
    
    // Clean up input postcode (remove spaces, convert to uppercase)
    const cleanInput = inputPostcode.replace(/\s+/g, '').toUpperCase();
    
    // Split coverage areas by semicolon or comma (e.g., "G1-80;EH1-55;FK1-20" or "G12, EH6, FK15")
    const areas = coverageString.split(/[;,]/).map(area => area.trim()).filter(area => area.length > 0);
    
    // First check for exact matches (including partial matches)
    const exactMatch = areas.some(area => {
      return cleanInput === area.toUpperCase();
    });
    
    if (exactMatch) {
      return true;
    }
    
    // Check for partial matches (e.g., "G1" should match "G1-80")
    const partialMatch = areas.some(area => {
      return area.toUpperCase().startsWith(cleanInput);
    });
    
    if (partialMatch) {
      return true;
    }
    
    const result = areas.some(area => {
      
      // First try to parse range format like "G1-80" 
      const rangeMatch = area.match(/^([A-Z]+)(\d+)-(\d+)$/);
      if (rangeMatch) {
        const [, prefix, startNum, endNum] = rangeMatch;
        const start = parseInt(startNum);
        const end = parseInt(endNum);
        
        // Check if input matches the prefix and is within range
        const inputMatch = cleanInput.match(/^([A-Z]+)(\d+)$/);
        if (!inputMatch) {
          // If input doesn't match single postcode format, try exact match with the range
          const exactMatch = cleanInput === area.toUpperCase();
          return exactMatch;
        }
        
        const [, inputPrefix, inputNumStr] = inputMatch;
        const inputNum = parseInt(inputNumStr);
        
        const inRange = inputPrefix === prefix && inputNum >= start && inputNum <= end;
        return inRange;
      }
      
      // If not a range, try exact match (e.g., "G12", "EH6", "FK15")
      const exactMatch = area.match(/^([A-Z]+)(\d+)$/);
      if (exactMatch) {
        const exactResult = cleanInput === area.toUpperCase();
        return exactResult;
      }
      
      return false;
    });
    
    return result;
  }
  
  function populateFilterOptions() {
    console.log('=== POPULATING FILTER OPTIONS ===');
    console.log('Fitters array length:', fitters.length);
    console.log('Orders array length:', orders.length);
    console.log('Sample fitter data:', fitters.slice(0, 2));
    console.log('Sample fitter postcodes:', fitters.slice(0, 3).map(f => ({ name: f.name, postcodesCovered: f.postcodesCovered })));
    
    // Fitter filter - setup searchable select
    const fitterFilter = document.getElementById('fitterFilter');
    const fitterSearch = document.getElementById('fitterSearch');
    const fitterDropdown = document.getElementById('fitterDropdown');
    
    if (!fitterFilter || !fitterSearch || !fitterDropdown) {
      console.error('Fitter filter elements not found!');
      return;
    }
    
     // Clear existing options first
    fitterFilter.innerHTML = '<option value="">All Fitters</option>';
     fitterDropdown.innerHTML = '';
     
     // Get unique fitters and sort alphabetically
     const uniqueFitters = [...new Set(fitters.map(f => f.name))].sort((a, b) => a.localeCompare(b));
     
     // Populate both select and dropdown
     uniqueFitters.forEach(fitterName => {
      const option = document.createElement('option');
       option.value = fitterName;
       option.textContent = fitterName;
      fitterFilter.appendChild(option);
       
       const dropdownOption = document.createElement('div');
       dropdownOption.className = 'select-option';
       dropdownOption.textContent = fitterName;
       dropdownOption.onclick = () => selectFitter(fitterName);
       fitterDropdown.appendChild(dropdownOption);
     });
    
    // Remove existing event listeners to prevent duplicates
    fitterSearch.removeEventListener('input', handleFitterSearch);
    fitterSearch.removeEventListener('focus', handleFitterFocus);
    fitterSearch.removeEventListener('blur', handleFitterBlur);
    
    // Define event handlers
    function handleFitterSearch(e) {
      const searchTerm = e.target.value.toLowerCase();
      const options = fitterDropdown.querySelectorAll('.select-option');
      
      options.forEach(option => {
        const text = option.textContent.toLowerCase();
        if (text.includes(searchTerm)) {
          option.style.display = 'block';
        } else {
          option.style.display = 'none';
        }
      });
      
      fitterDropdown.classList.add('show');
      
      // If search is cleared, show all options and reset filter
      if (searchTerm === '') {
        options.forEach(option => {
          option.style.display = 'block';
        });
        fitterFilter.value = '';
        handleFilterChange();
      }
    }
    
    function handleFitterFocus() {
      fitterDropdown.classList.add('show');
    }
    
    function handleFitterBlur() {
      setTimeout(() => {
        fitterDropdown.classList.remove('show');
      }, 200);
    }
    
    // Setup search functionality
    fitterSearch.addEventListener('input', handleFitterSearch);
    fitterSearch.addEventListener('focus', handleFitterFocus);
    fitterSearch.addEventListener('blur', handleFitterBlur);
    
    console.log('Fitter filter populated with', fitters.length, 'options');
    
    // Customer filter - setup searchable select
    const customerFilter = document.getElementById('customerFilter');
    const customerSearch = document.getElementById('customerSearch');
    const customerDropdown = document.getElementById('customerDropdown');
    
    if (!customerFilter || !customerSearch || !customerDropdown) {
      console.error('Customer filter elements not found!');
      return;
    }
    
    // Clear existing options first
    customerFilter.innerHTML = '<option value="">All Customers</option>';
    customerDropdown.innerHTML = '';
    
    const customers = [...new Set(orders.map(order => {
      // Handle both string and object customer names
      if (typeof order.customerName === 'string') {
        return order.customerName;
      } else if (order.customerName && order.customerName.name) {
        return order.customerName.name;
      }
      return null;
    }).filter(name => name && 
                      name !== 'Unknown Customer' && 
                      !name.startsWith('Customer ') &&
                      name.length > 0))];
    
    console.log('Customer filter options:', customers);
    
     // Ensure unique customers and sort alphabetically
     const uniqueCustomers = [...new Set(customers)].sort((a, b) => a.localeCompare(b));
     
     // Populate both select and dropdown
     uniqueCustomers.forEach(customer => {
      const option = document.createElement('option');
      option.value = customer;
      option.textContent = customer;
      customerFilter.appendChild(option);
       
       const dropdownOption = document.createElement('div');
       dropdownOption.className = 'select-option';
       dropdownOption.textContent = customer;
       dropdownOption.onclick = () => selectCustomer(customer);
       customerDropdown.appendChild(dropdownOption);
     });
    
    // Remove existing event listeners to prevent duplicates
    customerSearch.removeEventListener('input', handleCustomerSearch);
    customerSearch.removeEventListener('focus', handleCustomerFocus);
    customerSearch.removeEventListener('blur', handleCustomerBlur);
    
    // Define event handlers
    function handleCustomerSearch(e) {
      const searchTerm = e.target.value.toLowerCase();
      const options = customerDropdown.querySelectorAll('.select-option');
      
      options.forEach(option => {
        const text = option.textContent.toLowerCase();
        if (text.includes(searchTerm)) {
          option.style.display = 'block';
        } else {
          option.style.display = 'none';
        }
      });
      
      customerDropdown.classList.add('show');
      
      // If search is cleared, show all options and reset filter
      if (searchTerm === '') {
        options.forEach(option => {
          option.style.display = 'block';
        });
        customerFilter.value = '';
        handleFilterChange();
      }
    }
    
    function handleCustomerFocus() {
      customerDropdown.classList.add('show');
    }
    
    function handleCustomerBlur() {
      setTimeout(() => {
        customerDropdown.classList.remove('show');
      }, 200);
    }
    
    // Setup search functionality
    customerSearch.addEventListener('input', handleCustomerSearch);
    customerSearch.addEventListener('focus', handleCustomerFocus);
    customerSearch.addEventListener('blur', handleCustomerBlur);
    
    // Skillset filter - get all unique skillsets from all fitters
    const skillsetFilter = document.getElementById('skillsetFilter');
    if (!skillsetFilter) {
      console.error('Skillset filter element not found!');
      return;
    }
    
    skillsetFilter.innerHTML = '<option value="">All Skillsets</option>';
    
    console.log('=== SKILLSET FILTER DEBUG ===');
    console.log('Fitters for skillset filter:', fitters.length);
    console.log('Sample fitter skillsets:', fitters.slice(0, 3).map(f => ({ name: f.name, skillset: f.skillset })));
    
    const allSkillsets = [...new Set(fitters.flatMap(f => {
      // Use the correct field name from Zoho CRM
      const skillsetField = f.skillset; // Already mapped in processFittersData
      console.log(`Fitter ${f.name} skillset: "${skillsetField}" (type: ${typeof skillsetField})`);
      
      if (!skillsetField) {
        console.log(`Skipping fitter ${f.name} - no skillset data`);
        return [];
      }
      
      let skills = [];
      if (Array.isArray(skillsetField)) {
        // If it's already an array, use it directly
        skills = skillsetField.map(skill => skill.trim()).filter(skill => skill.length > 0);
      } else if (typeof skillsetField === 'string') {
        // If it's a string, split it
        skills = skillsetField.split(/[;,]/).map(skill => skill.trim()).filter(skill => skill.length > 0);
      } else {
        console.log(`Skipping fitter ${f.name} - invalid skillset type: ${typeof skillsetField}`);
        return [];
      }
      
      console.log(`Fitter ${f.name} parsed skills:`, skills);
      return skills;
    }))];
    
    console.log('All skillsets found:', allSkillsets);
    console.log('Skillset filter element:', skillsetFilter);
    
    // Sort skillsets alphabetically
    const sortedSkillsets = allSkillsets.sort((a, b) => a.localeCompare(b));
    
    sortedSkillsets.forEach(skillset => {
      const option = document.createElement('option');
      option.value = skillset;
      option.textContent = skillset;
      skillsetFilter.appendChild(option);
    });
    
    console.log('Skillset filter populated with', allSkillsets.length, 'options');
    
    // Skills Shortage filter - get all unique skills shortages from all fitters
    const skillsShortageDropdown = document.getElementById('skillsShortageDropdown');
    if (!skillsShortageDropdown) {
      console.error('❌ Skills Shortage dropdown element not found!');
      return;
    }
    
    console.log('✅ Skills Shortage dropdown element found:', skillsShortageDropdown);
    skillsShortageDropdown.innerHTML = '';
    
    console.log('=== SKILLS SHORTAGE FILTER DEBUG ===');
    console.log('Fitters for skills shortage filter:', fitters.length);
    console.log('Sample fitter skills shortages:', fitters.slice(0, 3).map(f => ({ name: f.name, skillsShortage: f.skillsShortage })));
    
    const allSkillsShortages = [...new Set(fitters.flatMap(f => {
      const skillsShortageField = f.skillsShortage; // Already mapped in processFittersData
      console.log(`Fitter ${f.name} skills shortage:`, skillsShortageField, `(type: ${typeof skillsShortageField})`);
      
      if (!skillsShortageField || skillsShortageField === '') {
        console.log(`Skipping fitter ${f.name} - no skills shortage data`);
        return [];
      }
      
      let shortages = [];
      if (Array.isArray(skillsShortageField)) {
        // If it's already an array, process each element
        console.log(`Fitter ${f.name} skills shortage array:`, skillsShortageField);
        shortages = skillsShortageField.flatMap(item => {
          if (typeof item === 'string') {
            return parseSkillsShortageString(item);
          }
          return [item];
        }).filter(shortage => shortage.length > 0);
      } else if (typeof skillsShortageField === 'string') {
        // If it's a string, split it
        shortages = skillsShortageField.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
      } else if (typeof skillsShortageField === 'object' && skillsShortageField !== null) {
        // If it's an object, try to extract the value
        console.log(`Fitter ${f.name} skills shortage object:`, skillsShortageField);
        console.log(`Object keys:`, Object.keys(skillsShortageField));
        console.log(`Object values:`, Object.values(skillsShortageField));
        
        // Check if it has a value property or is a single-item object
        if (skillsShortageField.value) {
          console.log(`Using .value property:`, skillsShortageField.value);
          shortages = skillsShortageField.value.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
        } else if (Object.keys(skillsShortageField).length === 1) {
          // Single key object, use the value
          const value = Object.values(skillsShortageField)[0];
          console.log(`Using single key value:`, value);
          if (typeof value === 'string') {
            shortages = value.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
          }
        } else {
          // Try to find any string value in the object
          const stringValues = Object.values(skillsShortageField).filter(v => typeof v === 'string' && v.trim() !== '');
          if (stringValues.length > 0) {
            console.log(`Using first string value found:`, stringValues[0]);
            shortages = stringValues[0].split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
          } else {
            console.log(`Skipping fitter ${f.name} - no string values found in object:`, skillsShortageField);
            return [];
          }
        }
      } else {
        console.log(`Skipping fitter ${f.name} - invalid skills shortage type: ${typeof skillsShortageField}`);
        return [];
      }
      
      console.log(`Fitter ${f.name} parsed shortages:`, shortages);
      return shortages;
    }))];
    
    console.log('All skills shortages found:', allSkillsShortages);
    console.log('Unique skills shortages count:', allSkillsShortages.length);
    console.log('Skills shortages details:', allSkillsShortages.map(s => ({ value: s, length: s.length })));
    
    console.log('📝 Creating options for skills shortages...');
    allSkillsShortages.forEach((shortage, index) => {
      console.log(`Creating option ${index + 1}:`, shortage);
      const optionDiv = document.createElement('div');
      optionDiv.className = 'multi-select-option';
      
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.id = `skills-${shortage.replace(/[^a-zA-Z0-9]/g, '')}`;
      checkbox.value = shortage;
      checkbox.className = 'skills-shortage-checkbox';
      
      const label = document.createElement('label');
      label.htmlFor = checkbox.id;
      label.textContent = shortage;
      
      optionDiv.appendChild(checkbox);
      optionDiv.appendChild(label);
      skillsShortageDropdown.appendChild(optionDiv);
    });
    
    console.log(`✅ Created ${allSkillsShortages.length} options in dropdown`);
    console.log('Dropdown children count:', skillsShortageDropdown.children.length);
    
    // Initialize custom multi-select
    console.log('🔧 About to initialize custom multi-select...');
    initializeCustomMultiSelect();
    console.log('✅ Custom multi-select initialization completed');
    
    // Test: Try to manually open the dropdown
    setTimeout(() => {
      console.log('🧪 Testing manual dropdown open...');
      const testContainer = document.querySelector('.custom-multi-select');
      if (testContainer) {
        testContainer.classList.add('open');
        console.log('✅ Manual dropdown open test - class added');
        setTimeout(() => {
          testContainer.classList.remove('open');
          console.log('✅ Manual dropdown close test - class removed');
        }, 2000);
      } else {
        console.log('❌ Test container not found');
      }
      
      // Add inline onclick as a test
      const testTrigger = document.getElementById('skillsShortageTrigger');
      if (testTrigger) {
        testTrigger.onclick = function(e) {
          console.log('🖱️ Inline onclick triggered!');
          e.preventDefault();
          e.stopPropagation();
          const container = document.querySelector('.custom-multi-select');
          if (container) {
            container.classList.toggle('open');
            console.log('✅ Inline onclick - dropdown toggled');
          }
        };
        console.log('✅ Inline onclick handler added');
      } else {
        console.log('❌ Test trigger not found for inline onclick');
      }
    }, 1000);
    
    console.log('Skills Shortage filter populated with', allSkillsShortages.length, 'options');
    
    console.log(`Populated filters: ${fitters.length} fitters, ${customers.length} customers, ${allSkillsets.length} skillsets, ${allSkillsShortages.length} skills shortages`);
  }
  
  // =============================================
  // Holiday Data Fetching
  // =============================================
  
  async function loadHolidaysData() {
    console.log('=== LOADING HOLIDAYS DATA ===');
    
    if (!fitters || fitters.length === 0) {
      console.log('No fitters available for holiday data');
      return;
    }
    
    try {
      const holidayPromises = fitters.map(async (fitter) => {
        try {
          console.log(`Fetching holidays for fitter: ${fitter.name} (ID: ${fitter.id})`);
          
          const response = await ZOHO.CRM.API.getRelatedRecords({
            Entity: "Contacts",
            RecordID: fitter.id,
            RelatedList: "Fitter1"
          });
          
          if (response && response.data && response.data.length > 0) {
            console.log(`Found ${response.data.length} holiday records for ${fitter.name}`);
            
            return response.data.map(holiday => ({
              fitterId: fitter.id,
              fitterName: fitter.name,
              holidayId: holiday.id,
              title: holiday.Name || holiday.name || 'Holiday',
              status: holiday.Status || holiday.status || 'On Holiday',
              activeStatus: holiday.Active_Status || holiday.active_status || 'Booked',
              weekCommencing: holiday.Week_Commencing,
              createdTime: holiday.Created_Time,
              modifiedTime: holiday.Modified_Time
            }));
          } else {
            console.log(`No holiday records found for ${fitter.name}`);
            return [];
          }
        } catch (error) {
          console.error(`Error fetching holidays for fitter ${fitter.name}:`, error);
          return [];
        }
      });
      
      const allHolidays = await Promise.all(holidayPromises);
      const allHolidaysFlat = allHolidays.flat();
      
      // Filter to show only holidays with 'Booked' Active_Status
      holidays = allHolidaysFlat.filter(holiday => holiday.activeStatus === 'Booked');
      
      console.log(`Total holidays loaded: ${allHolidaysFlat.length}`);
      console.log(`Holidays with 'Booked' status: ${holidays.length}`);
      console.log('Sample holiday data:', holidays.slice(0, 3));
      
    } catch (error) {
      console.error('Error loading holidays data:', error);
    }
  }
  
  // Helper function to check if a date falls within a holiday week
  function isDateInHolidayWeek(date, fitterId) {
    if (!holidays || holidays.length === 0) {
      console.log('No holidays available for checking');
      return null;
    }
    
    const targetDate = new Date(date);
    const targetWeekStart = getWeekStart(targetDate);
    
    // Find holidays for this fitter
    const fitterHolidays = holidays.filter(holiday => holiday.fitterId === fitterId);
    console.log(`Checking holidays for fitter ${fitterId}:`, fitterHolidays);
    
    for (const holiday of fitterHolidays) {
      if (holiday.weekCommencing) {
        const holidayWeekStart = getWeekStart(new Date(holiday.weekCommencing));
        
        // Check if the target date falls within the holiday week
        if (targetWeekStart.getTime() === holidayWeekStart.getTime()) {
          return holiday;
        }
      }
    }
    
    return null;
  }
  
  // Helper function to get the start of the week (Monday)
  function getWeekStart(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1); // Adjust when day is Sunday
    return new Date(d.setDate(diff));
  }
  
  // =============================================
  // Calendar Display
  // =============================================
  
  function updateCalendar() {
  const weeks = getThreeWeekDays();
  const dateRange = getDateRange();
  
  // Update date range display with month names
  document.getElementById('dateRange').textContent = formatDateRange(dateRange.startDate, dateRange.endDate);
  
  // Update calendar header
  updateCalendarHeader(weeks);
  
  // Update calendar body
  updateCalendarBody(weeks);
  
  // Update today button state
  updateTodayButtonState();
  
  // Update counts
  updateCounts();
}
  
  function updateCalendarHeader(weeks) {
  const headerRow = document.querySelector('.calendar-header-row');
  
  // Clear existing week headers
  const existingHeaders = headerRow.querySelectorAll('.day-header');
  existingHeaders.forEach(header => header.remove());
  
  // Add week headers
  weeks.forEach(week => {
    const weekHeader = document.createElement('div');
    weekHeader.className = `day-header ${week.isToday ? 'today' : ''}`;
    weekHeader.innerHTML = `
      <div class="day-name">${week.label.split('\n')[0]}</div>
      <div class="day-date">${week.label.split('\n')[1]}</div>
    `;
    headerRow.appendChild(weekHeader);
  });
}
  
  function updateCalendarBody(weeks) {
  const calendarBody = document.getElementById('calendarBody');
  calendarBody.innerHTML = '';
  
  // Filter fitters based on current filters
  const filteredFitters = filterFitters();
  
  // Sort fitters with special order: Unassigned Fitter, Supply Only, then alphabetical
  const sortedFitters = filteredFitters.sort((a, b) => {
    // Special cases first
    if (a.name === 'Unassigned Fitter' && b.name !== 'Unassigned Fitter') return -1;
    if (b.name === 'Unassigned Fitter' && a.name !== 'Unassigned Fitter') return 1;
    if (a.name === 'Supply Only' && b.name !== 'Supply Only' && b.name !== 'Unassigned Fitter') return -1;
    if (b.name === 'Supply Only' && a.name !== 'Supply Only' && a.name !== 'Unassigned Fitter') return 1;
    
    // Then alphabetical order
    return a.name.localeCompare(b.name);
  });
  
  // Create fitter rows
  sortedFitters.forEach(fitter => {
    const fitterRow = createFitterRow(fitter, weeks);
    calendarBody.appendChild(fitterRow);
  });
}
  
  function createFitterRow(fitter, weeks) {
  const row = document.createElement('div');
  row.className = 'team-member-row';
  
  // Fitter info column
  const fitterInfo = document.createElement('div');
  fitterInfo.className = 'member-info';
  fitterInfo.innerHTML = `
    <div class="member-name">${fitter.name}</div>
    <div class="member-skills">${getDisplayValue(fitter.workConsistency)}</div>
  `;
  row.appendChild(fitterInfo);
  
  // Calendar week columns
  weeks.forEach(week => {
    const weekCell = createWeekCell(fitter, week);
    row.appendChild(weekCell);
  });
  
  return row;
}
  
  function createWeekCell(fitter, week) {
  const weekCell = document.createElement('div');
  weekCell.className = 'calendar-day';
  
  // Find orders for this fitter in this week that match current filters
  const weekOrders = orders.filter(order => {
    if (order.fitterId !== fitter.id) return false;
    
    // Check if order date falls within this week
    const orderDate = new Date(order.date);
    const weekStart = new Date(week.weekStart);
    const weekEnd = new Date(week.weekEnd);
    
    if (!(orderDate >= weekStart && orderDate <= weekEnd)) return false;
    
    // Apply current filters to orders
    // Filter by customer
    if (filters.customer) {
      const customerName = typeof order.customerName === 'string' ? order.customerName : 'Customer';
      if (!customerName.toLowerCase().includes(filters.customer.toLowerCase())) {
        return false;
      }
    }
    
    // Filter by job type
    if (filters.jobType) {
      if (!order.jobType.toLowerCase().includes(filters.jobType.toLowerCase())) {
        return false;
      }
    }
    
    // Note: Postcode filtering is handled at the fitter level, not order level
    // This ensures we only show orders from fitters whose coverage matches the filter
    
    return true;
  });
  
  // Sort orders by date and time
  weekOrders.sort((a, b) => {
    const dateCompare = a.date.localeCompare(b.date);
    if (dateCompare !== 0) return dateCompare;
    return a.startTime.localeCompare(b.startTime);
  });
  
  // Debug: Log filtered orders for this week
  if (weekOrders.length > 0) {
    console.log(`📅 Week ${week.weekNumber} - ${fitter.name}: ${weekOrders.length} filtered orders`);
    weekOrders.forEach(order => {
      console.log(`   Order: ${order.startTime} - ${order.customerName} (${order.jobType})`);
    });
  }
  
  // Check if it's current week
  if (week.isToday) {
    weekCell.classList.add('today');
  }
  
  // Check for holidays in this week
  const holiday = isDateInHolidayWeek(week.weekStart, fitter.id);
  
  // Add holiday tag first if there's a holiday in this week
  if (holiday) {
    const holidayElement = createHolidayElement(holiday);
    weekCell.appendChild(holidayElement);
  }
  
  // Add all orders to the week cell - stack times vertically
  if (weekOrders.length > 0) {
    weekOrders.forEach(order => {
      const orderElement = createOrderElement(order, fitter);
      weekCell.appendChild(orderElement);
    });
  } else if (!holiday) {
    // Show drop zone only if no orders AND no holiday
    const dropZone = document.createElement('div');
    dropZone.style.cssText = 'text-align: center; color: #9ca3b8; font-size: 0.75rem; padding: 1rem;';
    dropZone.textContent = '-';
    weekCell.appendChild(dropZone);
  }

  // Add click handler for creating holidays
  weekCell.addEventListener('click', (e) => {
    // Don't trigger if clicking on existing events or holidays
    if (e.target.classList.contains('calendar-event') || e.target.classList.contains('holiday')) {
      return;
    }
    
    // Only allow holiday creation if there's no existing holiday
    if (!holiday) {
      showCreateHolidayModal(fitter, week);
    }
  });
  
  return weekCell;
}
  
  // Helper function to convert text to proper case (first letter of each word capitalized)
  function toProperCase(text) {
    if (!text) return '';
    return text.toLowerCase().replace(/\b\w/g, l => l.toUpperCase());
  }

  function createHolidayElement(holiday) {
    console.log('createHolidayElement called with holiday:', holiday);
    const holidayElement = document.createElement('div');
    holidayElement.className = 'calendar-event holiday';
    
    // Get holiday title and convert to proper case, limit characters
    const holidayTitle = toProperCase(holiday.title || 'Holiday');
    const maxLength = 15; // Adjust based on column width
    const displayTitle = holidayTitle.length > maxLength 
      ? holidayTitle.substring(0, maxLength - 3) + '...' 
      : holidayTitle;
    
    holidayElement.textContent = displayTitle;
    holidayElement.title = `${toProperCase(holiday.fitterName)} - ${toProperCase(holiday.title || 'Holiday')} (Week of ${new Date(holiday.weekCommencing).toLocaleDateString()})`;
    holidayElement.onclick = (e) => {
      e.stopPropagation();
      console.log('Holiday element clicked:', holiday);
      showHolidayModal(holiday);
    };
    return holidayElement;
  }
  
  function createOrderElement(order, fitter) {
    const orderElement = document.createElement('div');
    orderElement.className = 'calendar-event';
    orderElement.onclick = async () => await showEventModal(order.id);
    
    // Determine order color based on job type
    const jobType = order.jobType?.toLowerCase() || '';
    if (jobType.includes('kitchen')) {
      orderElement.classList.add('green');
    } else if (jobType.includes('bathroom')) {
      orderElement.classList.add('blue');
    } else {
      orderElement.classList.add('black');
    }
    
    // Show customer name and order postcode in the calendar UI
    const customerNameStr = typeof order.customerName === 'string' ? order.customerName : 'Customer';
    
    // Use second part of customer name (e.g., "Fiona Quay" -> "Quay")
    const nameParts = customerNameStr.split(' ');
    const displayName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : customerNameStr;
    
    // Get postcode from order data and shorten it (first 3-4 characters before space)
    const orderPostcode = order.postcode || '';
    let shortPostcode = orderPostcode;
    if (orderPostcode !== '') {
      const spaceIndex = orderPostcode.indexOf(' ');
      if (spaceIndex > 0) {
        shortPostcode = orderPostcode.substring(0, spaceIndex);
      } else if (orderPostcode.length > 4) {
        shortPostcode = orderPostcode.substring(0, 4);
      }
    }
    
    orderElement.textContent = `${displayName} - ${shortPostcode}`;
    
    // Add title attribute for full order info on hover
    orderElement.title = `${order.startTime} - ${customerNameStr} (${order.jobType}) - ${orderPostcode}`;
    
    // Store the postcode in a data attribute for consistency
    orderElement.setAttribute('data-postcode', orderPostcode);
    
    return orderElement;
  }
  
  // =============================================
  // Filtering
  // =============================================
  
  function filterFitters() {
    
    return fitters.filter(fitter => {
      // Filter by fitter name - exact match for dropdown selection
      if (filters.fitter && fitter.name !== filters.fitter) {
        return false;
      }
      
      // Filter by skillset
      if (filters.skillset && (!fitter.skillset || !fitter.skillset.toLowerCase().includes(filters.skillset.toLowerCase()))) {
        return false;
      }
      
      // Filter by skills shortage - exclude fitters who have ANY of the selected skill shortages
      if (filters.skillsShortage && filters.skillsShortage.length > 0) {
        console.log(`🔍 Checking fitter ${fitter.name} for skill shortages:`, filters.skillsShortage);
        
        const fitterShortages = fitter.skillsShortage || '';
        let shortages = [];
        
        if (Array.isArray(fitterShortages)) {
          // Handle array data type - split comma-separated strings within array elements
          shortages = fitterShortages.flatMap(item => {
            if (typeof item === 'string') {
              return parseSkillsShortageString(item).map(s => s.toLowerCase());
            }
            return [item];
          }).filter(s => s.length > 0);
        } else if (typeof fitterShortages === 'string') {
          shortages = fitterShortages.split(/[;,]/).map(s => s.trim().toLowerCase()).filter(s => s.length > 0);
        } else if (typeof fitterShortages === 'object' && fitterShortages !== null) {
          // Handle object data type
          if (fitterShortages.value) {
            shortages = fitterShortages.value.split(/[;,]/).map(s => s.trim().toLowerCase()).filter(s => s.length > 0);
          } else if (Object.keys(fitterShortages).length === 1) {
            const value = Object.values(fitterShortages)[0];
            if (typeof value === 'string') {
              shortages = value.split(/[;,]/).map(s => s.trim().toLowerCase()).filter(s => s.length > 0);
            }
          }
        }
        
        console.log(`Fitter ${fitter.name} shortages:`, shortages);
        console.log(`Looking for any of:`, filters.skillsShortage.map(s => s.toLowerCase()));
        
        // Check if fitter has ANY of the selected skill shortages
        const hasAnySelectedShortage = filters.skillsShortage.some(selectedShortage => 
          shortages.some(shortage => 
            shortage.toLowerCase().includes(selectedShortage.toLowerCase()) || 
            selectedShortage.toLowerCase().includes(shortage.toLowerCase())
          )
        );
        
        if (hasAnySelectedShortage) {
          console.log(`❌ EXCLUDING fitter ${fitter.name} - HAS one or more selected skill shortages`);
          return false; // Exclude this fitter
        }
        
        console.log(`✅ INCLUDING fitter ${fitter.name} - does NOT have any selected skill shortages`);
      }
      
      // Filter by postcode - only show fitters whose postcode coverage matches
      if (filters.postcode) {
        const fitterPostcodeCoverage = fitter.postcodesCovered || '';
        const postcodeMatch = matchesPostcodeRange(filters.postcode, fitterPostcodeCoverage);
        if (!postcodeMatch) {
          return false;
        }
      }
      
      // Filter by customer - only show fitters who have orders with the selected customer
      if (filters.customer) {
        const fitterOrders = orders.filter(order => 
          order.fitterId === fitter.id && order.customerName === filters.customer
        );
        if (fitterOrders.length === 0) {
          return false;
        }
      }
      
      // Filter by job type - only show fitters who have orders of that type
      if (filters.jobType) {
        const fitterOrders = orders.filter(order => 
          order.fitterId === fitter.id && order.jobType === filters.jobType
        );
        if (fitterOrders.length === 0) {
          return false;
        }
      }
      
      return true;
    });
    
    const result = fitters.filter(fitter => {
      // Filter by fitter name - exact match for dropdown selection
      if (filters.fitter && fitter.name !== filters.fitter) {
        return false;
      }
      
      // Filter by skillset
      if (filters.skillset && (!fitter.skillset || !fitter.skillset.toLowerCase().includes(filters.skillset.toLowerCase()))) {
        return false;
      }
      
      // Filter by skills shortage - exclude fitters who have ANY of the selected skill shortages
      if (filters.skillsShortage && filters.skillsShortage.length > 0) {
        const fitterShortages = fitter.skillsShortage || '';
        let shortages = [];
        
        if (Array.isArray(fitterShortages)) {
          // Handle array data type - split comma-separated strings within array elements
          shortages = fitterShortages.flatMap(item => {
            if (typeof item === 'string') {
              return parseSkillsShortageString(item).map(s => s.toLowerCase());
            }
            return [item];
          }).filter(s => s.length > 0);
        } else if (typeof fitterShortages === 'string') {
          shortages = fitterShortages.split(/[;,]/).map(s => s.trim().toLowerCase()).filter(s => s.length > 0);
        } else if (typeof fitterShortages === 'object' && fitterShortages !== null) {
          // Handle object data type
          if (fitterShortages.value) {
            shortages = fitterShortages.value.split(/[;,]/).map(s => s.trim().toLowerCase()).filter(s => s.length > 0);
          } else if (Object.keys(fitterShortages).length === 1) {
            const value = Object.values(fitterShortages)[0];
            if (typeof value === 'string') {
              shortages = value.split(/[;,]/).map(s => s.trim().toLowerCase()).filter(s => s.length > 0);
            }
          }
        }
        
        // Check if fitter has ANY of the selected skill shortages
        const hasAnySelectedShortage = filters.skillsShortage.some(selectedShortage => 
          shortages.some(shortage => 
            shortage.toLowerCase().includes(selectedShortage.toLowerCase()) || 
            selectedShortage.toLowerCase().includes(shortage.toLowerCase())
          )
        );
        
        if (hasAnySelectedShortage) {
          return false;
        }
      }
      
      // Filter by postcode - only show fitters whose postcode coverage matches
      if (filters.postcode) {
        const fitterPostcodeCoverage = fitter.postcodesCovered || '';
        const postcodeMatch = matchesPostcodeRange(filters.postcode, fitterPostcodeCoverage);
        if (!postcodeMatch) {
          return false;
        }
      }
      
      // Filter by customer - only show fitters who have orders with the selected customer
      if (filters.customer) {
        const fitterOrders = orders.filter(order => 
          order.fitterId === fitter.id && order.customerName === filters.customer
        );
        if (fitterOrders.length === 0) {
          return false;
        }
      }
      
      return true;
    });
    
    return result;
  }
  
  function handleFilterChange() {
    filters.fitter = document.getElementById('fitterFilter').value;
    filters.customer = document.getElementById('customerFilter').value;
    filters.jobType = document.getElementById('jobTypeFilter').value;
    filters.skillset = document.getElementById('skillsetFilter').value;
    
    // Handle skills shortage multi-select
    const skillsShortageDropdown = document.getElementById('skillsShortageDropdown');
    if (skillsShortageDropdown) {
      const checkedBoxes = skillsShortageDropdown.querySelectorAll('input[type="checkbox"].skills-shortage-checkbox:checked');
      filters.skillsShortage = Array.from(checkedBoxes).map(checkbox => checkbox.value);
    } else {
      filters.skillsShortage = [];
    }
    
    filters.postcode = document.getElementById('postcodeFilter').value;
    
    updateCalendar();
  }
  
  function clearAllFilters() {
    filters = {
      fitter: '',
      customer: '',
      jobType: '',
      skillset: '',
      skillsShortage: [],
      postcode: ''
    };
    
    document.getElementById('fitterFilter').value = '';
    document.getElementById('customerFilter').value = '';
    document.getElementById('jobTypeFilter').value = '';
    document.getElementById('skillsetFilter').value = '';
    document.getElementById('postcodeFilter').value = '';
    
    // Clear search fields
    const fitterSearch = document.getElementById('fitterSearch');
    const customerSearch = document.getElementById('customerSearch');
    if (fitterSearch) fitterSearch.value = '';
    if (customerSearch) customerSearch.value = '';
    
    // Hide dropdowns
    const fitterDropdown = document.getElementById('fitterDropdown');
    const customerDropdown = document.getElementById('customerDropdown');
    if (fitterDropdown) fitterDropdown.classList.remove('show');
    if (customerDropdown) customerDropdown.classList.remove('show');
    
    // Clear skills shortage multi-select
    const skillsShortageDropdown = document.getElementById('skillsShortageDropdown');
    if (skillsShortageDropdown) {
      const checkboxes = skillsShortageDropdown.querySelectorAll('input[type="checkbox"].skills-shortage-checkbox');
      checkboxes.forEach(checkbox => checkbox.checked = false);
      
      // Update the display text
      const text = document.querySelector('.multi-select-text');
      if (text) {
        text.textContent = 'No Exclusions';
      }
    }
    
    updateCalendar();
  }
  
  // =============================================
  // User Name Fetching
  // =============================================
  
  async function fetchUserMemberName(userId) {
    if (!userId) {
      return 'Not assigned';
    }
    
    try {
      console.log(`Fetching user name for ID: ${userId}`);
      
      const response = await ZOHO.CRM.API.getUser({
        ID: userId
      });
      
      if (response && response.users && response.users.length > 0) {
        const user = response.users[0];
        
        // Use full_name directly, or fallback to first_name + last_name
        const fullName = user.full_name || `${user.first_name || ''} ${user.last_name || ''}`.trim();
        
        console.log(`User found: ${fullName}`);
        return fullName || 'No entry in Users';
      } else {
        console.log(`No user found for ID: ${userId}`);
        return 'No entry in Users';
      }
    } catch (error) {
      console.error(`Error fetching user ${userId}:`, error);
      return 'No entry in Users';
    }
  }

  // =============================================
  // Holiday Modal Functions
  // =============================================
  
  // Show create holiday modal
  function showCreateHolidayModal(fitter, week) {
    console.log('showCreateHolidayModal called with fitter:', fitter, 'week:', week);
    const modal = document.getElementById('createHolidayModal');
    const fitterNameInput = document.getElementById('holidayFitterName');
    const weekStartInput = document.getElementById('holidayWeekStart');
    const titleInput = document.getElementById('holidayTitle');
    
    if (!modal || !fitterNameInput || !weekStartInput || !titleInput) {
      console.error('Create holiday modal elements not found!');
      return;
    }
    
    // Pre-fill the form
    fitterNameInput.value = fitter.name;
    
    // Format the week start date
    const weekStartDate = new Date(week.weekStart);
    const formattedDate = weekStartDate.toLocaleDateString('en-GB', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    weekStartInput.value = formattedDate;
    
    // Clear the title input
    titleInput.value = '';
    
    // Update character count
    updateCharacterCount();
    
    // Add character count listener
    titleInput.addEventListener('input', updateCharacterCount);
    
    // Store the fitter and week data for later use
    modal.dataset.fitterId = fitter.id;
    modal.dataset.fitterName = fitter.name;
    modal.dataset.weekStart = week.weekStart;
    
    // Show modal
    modal.style.display = 'block';
  }
  
  // Create holiday record
  async function createHolidayRecord(fitterId, fitterName, weekStart, holidayTitle) {
    try {
      console.log('Creating holiday record:', { fitterId, fitterName, weekStart, holidayTitle });
      
      const recordData = {
        Name: holidayTitle,
        Fitter: fitterId,
        Week_Commencing: weekStart,
        Active_Status: 'Booked'
      };
      
      console.log('Record data to be sent:', recordData);
      
      const response = await ZOHO.CRM.API.insertRecord({
        Entity: "Fitter_Unavailability",
        APIData: recordData
      });
      
      console.log('Holiday creation response:', response);
      
      if (response && response.data && response.data.length > 0) {
        const result = response.data[0];
        if (result.status === 'success') {
          console.log('Holiday created successfully:', result.details);
          return { success: true, holidayId: result.details.id };
        } else {
          console.error('Error creating holiday:', result.message);
          return { success: false, error: result.message };
        }
      } else {
        console.error('Unexpected response format:', response);
        return { success: false, error: 'Unexpected response format' };
      }
    } catch (error) {
      console.error('Error creating holiday record:', error);
      return { success: false, error: error.message };
    }
  }
  
  // Update character count for holiday title
  function updateCharacterCount() {
    const titleInput = document.getElementById('holidayTitle');
    const countDisplay = document.getElementById('holidayTitleCount');
    
    if (titleInput && countDisplay) {
      const currentLength = titleInput.value.length;
      const maxLength = 30;
      
      countDisplay.textContent = `${currentLength}/${maxLength} characters`;
      
      // Change color based on character count
      if (currentLength > maxLength * 0.9) {
        countDisplay.style.color = '#dc2626'; // Red when close to limit
      } else if (currentLength > maxLength * 0.7) {
        countDisplay.style.color = '#f59e0b'; // Orange when getting close
      } else {
        countDisplay.style.color = '#6b7280'; // Gray for normal
      }
    }
  }
  
  // Close create holiday modal
  function closeCreateHolidayModal() {
    const modal = document.getElementById('createHolidayModal');
    if (modal) {
      modal.style.display = 'none';
    }
  }
  
  // Handle create holiday form submission
  async function handleCreateHolidaySubmit() {
    const modal = document.getElementById('createHolidayModal');
    const titleInput = document.getElementById('holidayTitle');
    
    if (!modal || !titleInput) {
      console.error('Modal or title input not found');
      return;
    }
    
    const holidayTitle = titleInput.value.trim();
    if (!holidayTitle) {
      alert('Please enter a holiday title');
      return;
    }
    
    if (holidayTitle.length > 30) {
      alert('Holiday title must be 30 characters or less');
      return;
    }
    
    const fitterId = modal.dataset.fitterId;
    const fitterName = modal.dataset.fitterName;
    const weekStart = modal.dataset.weekStart;
    
    if (!fitterId || !weekStart) {
      console.error('Missing required data for holiday creation');
      alert('Error: Missing required data');
      return;
    }
    
    // Show loading state
    const submitBtn = document.getElementById('saveCreateHoliday');
    const originalText = submitBtn.textContent;
    submitBtn.textContent = 'Creating...';
    submitBtn.disabled = true;
    
    try {
      const result = await createHolidayRecord(fitterId, fitterName, weekStart, holidayTitle);
      
      if (result.success) {
        console.log('Holiday created successfully');
        alert('Holiday created successfully!');
        closeCreateHolidayModal();
        
        // Refresh the calendar to show the new holiday
        await refreshCalendarData();
      } else {
        console.error('Failed to create holiday:', result.error);
        alert('Error creating holiday: ' + result.error);
      }
    } catch (error) {
      console.error('Error in holiday creation process:', error);
      alert('Error creating holiday: ' + error.message);
    } finally {
      // Restore button state
      submitBtn.textContent = originalText;
      submitBtn.disabled = false;
    }
  }
  
  // Refresh calendar data after holiday creation
  async function refreshCalendarData() {
    console.log('Refreshing calendar data...');
    try {
      // Reload holidays data
      await loadHolidaysData();
      
      // Re-render the calendar (this will maintain current filters and view)
      updateCalendar();
      
      console.log('Calendar data refreshed successfully');
    } catch (error) {
      console.error('Error refreshing calendar data:', error);
    }
  }
  
  function showHolidayModal(holiday) {
    console.log('showHolidayModal called with holiday:', holiday);
    const modal = document.getElementById('eventModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalBody = document.getElementById('modalBody');
    
    if (!modal || !modalTitle || !modalBody) {
      console.error('Holiday modal elements not found!', {modal, modalTitle, modalBody});
      return;
    }
    
    // Set modal title with edit and delete buttons
    modalTitle.innerHTML = `
      <span>${toProperCase(holiday.title || 'Holiday Details')}</span>
      <div style="display: flex; gap: 0.5rem; margin-left: auto;">
        <a href="https://crm.zoho.com/crm/org2579410/tab/CustomModule8/${holiday.holidayId}/edit?layoutId=167246000052041066" 
           target="_blank" 
           class="btn btn-edit" 
           style="background-color: #3b82f6; color: white; padding: 0.5rem 1rem; border-radius: 4px; text-decoration: none; display: inline-block; font-size: 0.875rem;">
          Edit Holiday
        </a>
        <button onclick="deleteHoliday('${holiday.holidayId}', '${holiday.fitterId}')" 
                class="btn btn-delete" 
                style="background-color: #dc2626; color: white; padding: 0.5rem 1rem; border-radius: 4px; border: none; cursor: pointer; font-size: 0.875rem; display: flex; align-items: center; gap: 0.25rem;">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="3,6 5,6 21,6"></polyline>
            <path d="m19,6v14a2,2 0 0,1 -2,2H7a2,2 0 0,1 -2,-2V6m3,0V4a2,2 0 0,1 2,-2h4a2,2 0 0,1 2,2v2"></path>
            <line x1="10" y1="11" x2="10" y2="17"></line>
            <line x1="14" y1="11" x2="14" y2="17"></line>
          </svg>
          Delete
        </button>
      </div>
`;
    
    // Format the holiday date
    const holidayDate = new Date(holiday.weekCommencing);
    const formattedDate = holidayDate.toLocaleDateString('en-GB', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    
    // Calculate the end of the holiday week (Sunday)
    const weekEnd = new Date(holidayDate);
    weekEnd.setDate(weekEnd.getDate() + 6);
    const formattedWeekEnd = weekEnd.toLocaleDateString('en-GB', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    
    // Create modal content
    modalBody.innerHTML = `
      <div class="modal-field">
        <span class="modal-label">Fitter:</span>
        <span class="modal-value">
          <a href="${getContactUrl(holiday.fitterId)}" target="_blank" style="color: #007bff; text-decoration: underline;">
            ${toProperCase(holiday.fitterName)}
          </a>
        </span>
      </div>
      <div class="modal-field">
        <span class="modal-label">Holiday Period:</span>
        <span class="modal-value">${formattedDate} - ${formattedWeekEnd}</span>
      </div>
        `;
    
    // Show modal
    console.log('Showing holiday modal with content:', modalBody.innerHTML);
    modal.style.display = 'block';
  }
  
  // Delete holiday function - called from onclick handler
  async function deleteHoliday(holidayId, fitterId) {
    console.log('deleteHoliday called with:', { holidayId, fitterId });
    
    // Show confirmation dialog
    const confirmed = confirm('Are you sure you want to delete this holiday? This will make the holiday \'Inactive\' and remove it from the calendar view');
    if (!confirmed) {
      return;
    }
    
    try {
      // Show loading state
      const modal = document.getElementById('eventModal');
      if (modal) {
        modal.style.display = 'none'; // Close modal while processing
      }
      
      showLoading(true, 'Deleting holiday...');
      
      // Update the holiday record to mark Active_Status as 'Cancelled'
      const config = {
        Entity: "Fitter_Unavailability",
        APIData: {
          id: holidayId,
          Active_Status: 'Cancelled'
        }
      };
      
      console.log('Updating holiday record with config:', config);
      
      const response = await ZOHO.CRM.API.updateRecord(config);
      
      if (response && response.data && response.data.length > 0) {
        const result = response.data[0];
        if (result.status === 'success') {
          console.log('Holiday successfully cancelled:', result);
          
          // Remove the holiday from the local array
          holidays = holidays.filter(holiday => holiday.holidayId !== holidayId);
          
          // Refresh the calendar to reflect changes
          await refreshCalendarData();
          
          alert('Holiday has been successfully cancelled and removed from the calendar.');
        } else {
          console.error('Failed to cancel holiday:', result);
          alert('Failed to delete holiday. Please try again.');
        }
      } else {
        console.error('No response data received:', response);
        alert('Failed to delete holiday. Please try again.');
      }
      
    } catch (error) {
      console.error('Error deleting holiday:', error);
      alert('An error occurred while deleting the holiday. Please try again.');
    } finally {
      showLoading(false);
    }
  }
  
  // Make deleteHoliday function globally available
  window.deleteHoliday = deleteHoliday;
  
  // =============================================
  // Event Modal
  // =============================================
  
  async function showEventModal(orderId) {
    console.log('showEventModal called with orderId:', orderId);
    const order = orders.find(o => o.id === orderId);
    if (!order) {
      console.error('Order not found for ID:', orderId);
      return;
    }
    
    // Debug: Log the order data to see what we're working with
    console.log('Order data for modal:', order);
    
    // Find the fitter to get additional details
    const fitter = fitters.find(f => f.id === order.fitterId);
    
    const modal = document.getElementById('eventModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalBody = document.getElementById('modalBody');
    
    // Set modal title
    modalTitle.textContent = 'Project Details';
    
    // Show loading state while fetching team member names
    modalBody.innerHTML = `
      <div class="modal-section">
          <div class="modal-field">
          <span class="modal-label">Title:</span>
          <span class="modal-value">${order.customerName} - ${getDisplayValue(order.postcode)}</span>
          </div>
        <div class="modal-field">
          <span class="modal-label">Customer:</span>
          <span class="modal-value">
            <a href="${getOrderUrl(order.id)}" target="_blank" style="color: #007bff; text-decoration: underline;">
              ${order.customerName}
            </a>
          </span>
        </div>
        <div class="modal-field">
          <span class="modal-label">Address:</span>
          <span class="modal-value">${getDisplayValue(order.address)}</span>
        </div>
        <div class="modal-field">
          <span class="modal-label">Postcode:</span>
          <span class="modal-value">${getDisplayValue(order.postcode)}</span>
        </div>
        <div class="modal-field">
          <span class="modal-label">Mobile:</span>
          <span class="modal-value">${getDisplayValue(order.mobile)}</span>
        </div>
      </div>
      
      <div class="modal-divider"></div>
      
      <div class="modal-section">
      <div class="modal-field">
          <span class="modal-label">Project Type:</span>
          <span class="modal-value">${getDisplayValue(order.jobType)}</span>
      </div>
      <div class="modal-field">
          <span class="modal-label">Showroom:</span>
          <span class="modal-value">${getDisplayValue(order.showroom)}</span>
      </div>
      <div class="modal-field">
          <span class="modal-label">Designer:</span>
          <span class="modal-value loading">Loading user...</span>
      </div>
      <div class="modal-field">
          <span class="modal-label">Surveyor:</span>
          <span class="modal-value loading">Loading user...</span>
      </div>
      <div class="modal-field">
          <span class="modal-label">Goods:</span>
          <span class="modal-value">${formatCurrency(order.goods)}</span>
        </div>
        <div class="modal-field">
          <span class="modal-label">Install:</span>
          <span class="modal-value">${formatCurrency(order.install)}</span>
        </div>
      </div>
    `;
    
    console.log('Showing order modal with content:', modalBody.innerHTML);
    modal.style.display = 'block';
    
    // Fetch user names asynchronously
    try {
      const [designerName, surveyorName] = await Promise.all([
        fetchUserMemberName(order.designer),
        fetchUserMemberName(order.surveyor)
      ]);
      
      // Update the modal with the fetched names
      // Find the Designer and Surveyor fields by looking for their labels
      const modalFields = modalBody.querySelectorAll('.modal-field');
      let designerField = null;
      let surveyorField = null;
      
      modalFields.forEach(field => {
        const label = field.querySelector('.modal-label');
        if (label && label.textContent === 'Designer:') {
          designerField = field.querySelector('.modal-value');
        } else if (label && label.textContent === 'Surveyor:') {
          surveyorField = field.querySelector('.modal-value');
        }
      });
      
      console.log('Designer field found:', designerField);
      console.log('Surveyor field found:', surveyorField);
      
      if (designerField) {
        designerField.textContent = designerName;
        designerField.classList.remove('loading');
        console.log('Updated designer field with:', designerName);
      }
      if (surveyorField) {
        surveyorField.textContent = surveyorName;
        surveyorField.classList.remove('loading');
        console.log('Updated surveyor field with:', surveyorName);
      }
      
    } catch (error) {
      console.error('Error fetching user names:', error);
      
      // Update with error state
      const modalFields = modalBody.querySelectorAll('.modal-field');
      let designerField = null;
      let surveyorField = null;
      
      modalFields.forEach(field => {
        const label = field.querySelector('.modal-label');
        if (label && label.textContent === 'Designer:') {
          designerField = field.querySelector('.modal-value');
        } else if (label && label.textContent === 'Surveyor:') {
          surveyorField = field.querySelector('.modal-value');
        }
      });
      
      if (designerField) {
        designerField.textContent = 'Error loading name';
        designerField.classList.remove('loading');
        designerField.classList.add('error');
      }
      if (surveyorField) {
        surveyorField.textContent = 'Error loading name';
        surveyorField.classList.remove('loading');
        surveyorField.classList.add('error');
      }
    }
  }
  
  function closeEventModal() {
    console.log('closeEventModal called');
    const modal = document.getElementById('eventModal');
    if (modal) {
      console.log('Closing modal');
      modal.style.display = 'none';
      modal.classList.remove('show'); // Remove any show class if it exists
    } else {
      console.error('Modal not found for closing');
    }
  }
  
  // =============================================
  // Utility Functions
  // =============================================
  
  function updateCounts() {
    const filteredFitters = filterFitters();
    const filteredOrders = orders.filter(order => {
      const fitter = fitters.find(f => f.id === order.fitterId);
      if (!fitter) return false;
      
      // Apply filters
      if (filters.fitter && !fitter.name.toLowerCase().includes(filters.fitter.toLowerCase())) {
        return false;
      }
      if (filters.customer && order.customerName !== filters.customer) {
        return false;
      }
      if (filters.jobType && order.jobType !== filters.jobType) {
        return false;
      }
      if (filters.skillset && (!fitter.skillset || !fitter.skillset.toLowerCase().includes(filters.skillset.toLowerCase()))) {
        return false;
      }
      // Note: Postcode filtering is handled at the fitter level in filterFitters()
      // This ensures we only count orders from fitters whose coverage matches the filter
      
      return true;
    });
    
    document.getElementById('eventCount').textContent = `${filteredOrders.length} orders`;
  }
  
  function showLoading(show) {
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.style.display = show ? 'flex' : 'none';
  }
  
  function showError(message) {
    console.error('Calendar Widget Error:', message);
    
    // Hide loading overlay
    showLoading(false);
    
    // Show error message in the calendar body
    const calendarBody = document.getElementById('calendarBody');
    if (calendarBody) {
      calendarBody.innerHTML = `
        <div style="text-align: center; padding: 2rem; color: #dc2626; background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; margin: 1rem;">
          <h3 style="margin-bottom: 1rem;">⚠️ Error Loading Calendar</h3>
          <p style="margin-bottom: 1rem;">${message}</p>
          <p style="font-size: 0.9rem; color: #6b7280;">Please contact the administrator or try refreshing the page.</p>
        </div>
      `;
    }
    
    // Update counts to show error state
    document.getElementById('eventCount').textContent = 'Error loading data';
  }
  
  function showNoEventsMessage() {
    console.log('No orders found for the current date range');
    
    // Show no events message in the calendar body
    const calendarBody = document.getElementById('calendarBody');
    if (calendarBody) {
      calendarBody.innerHTML = `
        <div style="text-align: center; padding: 2rem; color: #6b7280; background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; margin: 1rem;">
          <h3 style="margin-bottom: 1rem;">📅 No Orders Found</h3>
          <p style="margin-bottom: 1rem;">No orders found for the selected date range.</p>
          <p style="font-size: 0.9rem; color: #9ca3af;">Try selecting a different date range or contact the administrator if this seems incorrect.</p>
        </div>
      `;
    }
    
    // Update counts to show no events
    document.getElementById('eventCount').textContent = '0 orders';
  }
  
  // =============================================
  // View Button Navigation Functionality
  // =============================================
  
  function initializeTabs() {
    const viewButtons = document.querySelectorAll('.view-btn');
    const viewContents = document.querySelectorAll('.view-content');
    
    viewButtons.forEach(button => {
      button.addEventListener('click', () => {
        const targetView = button.getAttribute('data-view');
        
        // Remove active class from all buttons and contents
        viewButtons.forEach(btn => btn.classList.remove('active'));
        viewContents.forEach(content => content.classList.remove('active'));
        
        // Add active class to clicked button and corresponding content
        button.classList.add('active');
        const targetContent = document.getElementById(targetView + 'View');
        if (targetContent) {
          targetContent.classList.add('active');
        }
        
        // Load fitters data when fitters view is clicked
        if (targetView === 'fitters') {
          loadFittersData();
        }
      });
    });
  }
  
  // =============================================
  // Fitters Tab Functionality
  // =============================================
  
  function loadFittersData() {
    const fittersCount = document.getElementById('fittersCount');
    const fittersTableBody = document.getElementById('fittersTableBody');
    
    if (!fittersCount || !fittersTableBody) {
      console.error('Fitters view elements not found!');
      return;
    }
    
    // Use existing fitter data from calendar loading
    if (window.fitters && window.fitters.length > 0) {
      console.log('Using existing fitter data:', window.fitters.length, 'fitters');
      populateFittersTable(window.fitters);
      
      // Add event listeners for filters
      const fitterNameFilter = document.getElementById('fitterNameFilter');
      const showIncompleteOnly = document.getElementById('showIncompleteOnly');
      
      // Note: The fitter name filter is handled in populateFitterDropdown
      // via click events on dropdown items, so we don't need a change listener here
      
      if (showIncompleteOnly) {
        showIncompleteOnly.addEventListener('change', () => {
          populateFittersTable(window.fitters);
        });
      }
    } else {
      console.log('No existing fitter data available, showing message');
      fittersCount.textContent = 'No fitters data available';
      fittersTableBody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 2rem; color: #6b7280;">No fitters data available. Please load the calendar first.</td></tr>';
    }
  }
  
  async function fetchFittersData() {
    try {
      const response = await ZOHO.CRM.API.coql({
        Entity: "Contacts",
        cvid: "5000000000000000000", // Replace with your custom view ID
        select_query: "select First_Name,Last_Name,Email,Home_Phone,Mobile,Fitter_Skillset_2,Skills_They_Can_t_Do,Postcode_Area_New,Work_Consistency,Insurance_Status,PL_Date,PL_Insurance_Number,Terms_Masdouk,Fitter_Status from Contacts where Fitter_Status = 'Live'"
      });
      
      if (response && response.data) {
        return response.data.map(contact => ({
          id: contact.id,
          firstName: contact.First_Name || '',
          lastName: contact.Last_Name || '',
          email: contact.Email || '',
          phone: contact.Home_Phone || contact.Mobile || '',
          skillset: contact.Fitter_Skillset_2 || [],
          skillsShortage: contact.Skills_They_Can_t_Do || [],
          postcodesCovered: contact.Postcode_Area_New || '',
          workConsistency: contact.Work_Consistency || '',
          insuranceStatus: contact.Insurance_Status || '',
          plDate: contact.PL_Date || '',
          plInsuranceNo: contact.PL_Insurance_Number || '',
          termsMasdouk: contact.Terms_Masdouk || '',
          fitterStatus: contact.Fitter_Status || ''
        }));
      }
      return [];
    } catch (error) {
      console.error('Error fetching fitters:', error);
      throw error;
    }
  }
  
  // Calculate audit status for a fitter
  function calculateAuditStatus(fitter) {
    // Check if any of the required fields are missing
    // Fitter Skillset
    let hasSkillset = false;
    if (fitter.skillset) {
      if (Array.isArray(fitter.skillset)) {
        hasSkillset = fitter.skillset.length > 0;
      } else if (typeof fitter.skillset === 'string') {
        hasSkillset = fitter.skillset.trim() !== '' && fitter.skillset !== 'Missing Data';
      }
    }
    
    // Postcode Area
    const hasPostcodeArea = isDataAvailable(fitter.postcodesCovered);
    
    // Skills Shortage - this field should exist (even if empty array/string is valid)
    let hasSkillsShortage = false;
    if (fitter.skillsShortage !== undefined && fitter.skillsShortage !== null) {
      if (Array.isArray(fitter.skillsShortage)) {
        hasSkillsShortage = true; // Empty array is valid (means no shortages)
      } else if (typeof fitter.skillsShortage === 'string') {
        hasSkillsShortage = fitter.skillsShortage !== 'Missing Data';
      }
    }
    
    // Work Consistency
    const hasWorkConsistency = isDataAvailable(fitter.workConsistency);
    
    // Terms & Conditions
    const hasTermsConditions = isDataAvailable(fitter.termsMasdouk);
    
    // Insurance Status
    const hasInsuranceStatus = isDataAvailable(fitter.insuranceStatus);
    
    // Insurance Details (both PL Insurance Number and PL Date required)
    const hasInsuranceDetails = isDataAvailable(fitter.plInsuranceNo) && isDataAvailable(fitter.plDate);
    
    // If any field is missing, status is incomplete
    if (!hasSkillset || !hasPostcodeArea || !hasSkillsShortage || !hasWorkConsistency || 
        !hasTermsConditions || !hasInsuranceStatus || !hasInsuranceDetails) {
      return 'incomplete';
    }
    
    return 'complete';
  }

  // Populate fitter name dropdown (only called once or when fitters list changes)
  function populateFitterDropdown(fitters) {
    const dropdown = document.getElementById('fitterNameFilter');
    const input = document.getElementById('fitterNameFilterInput');
    const dropdownList = document.getElementById('fitterDropdownList');
    
    if (!dropdown || !input || !dropdownList) return;
    
    // Preserve current selection
    const currentSelection = dropdown.value;
    
    // Clear existing options
    dropdown.innerHTML = '<option value="">All Fitters</option>';
    dropdownList.innerHTML = '';
    
    // Sort fitters by name
    const sortedFitters = [...fitters].sort((a, b) => {
      const nameA = a.name.toLowerCase();
      const nameB = b.name.toLowerCase();
      return nameA.localeCompare(nameB);
    });
    
    // Add "All Fitters" option to dropdown list
    const allOption = document.createElement('div');
    allOption.className = 'fitter-dropdown-item';
    allOption.dataset.value = '';
    allOption.textContent = 'All Fitters';
    allOption.addEventListener('click', () => {
      input.value = 'All Fitters';
      dropdown.value = '';
      dropdownList.classList.remove('show');
      populateFittersTable(window.fitters);
    });
    dropdownList.appendChild(allOption);
    
    // Add each fitter as an option
    sortedFitters.forEach(fitter => {
      // Add to select dropdown (hidden)
      const option = document.createElement('option');
      option.value = fitter.id;
      option.textContent = fitter.name;
      dropdown.appendChild(option);
      
      // Add to custom dropdown list
      const item = document.createElement('div');
      item.className = 'fitter-dropdown-item';
      item.dataset.value = fitter.id;
      item.dataset.name = fitter.name;
      item.textContent = fitter.name;
      item.addEventListener('click', () => {
        input.value = fitter.name;
        dropdown.value = fitter.id;
        dropdownList.classList.remove('show');
        populateFittersTable(window.fitters);
      });
      dropdownList.appendChild(item);
    });
    
    // Restore previous selection
    if (currentSelection) {
      dropdown.value = currentSelection;
      const selectedFitter = fitters.find(f => f.id === currentSelection);
      if (selectedFitter) {
        input.value = selectedFitter.name;
      }
    } else {
      input.value = 'All Fitters';
    }
    
    // Add event listeners for searchable input (only if not already added)
    if (!input.hasAttribute('data-listeners-added')) {
      input.setAttribute('data-listeners-added', 'true');
      
      input.addEventListener('focus', () => {
        dropdownList.classList.add('show');
        filterDropdownList('');
      });
      
      input.addEventListener('input', (e) => {
        const searchTerm = e.target.value.toLowerCase();
        filterDropdownList(searchTerm);
        dropdownList.classList.add('show');
        
        // If input is cleared completely, reset to "All Fitters"
        if (e.target.value === '') {
          dropdown.value = '';
          populateFittersTable(window.fitters);
        }
      });
      
      input.addEventListener('blur', (e) => {
        // Delay hiding to allow click events
        setTimeout(() => {
          dropdownList.classList.remove('show');
        }, 200);
      });
    }
    
    // Filter dropdown list based on search term
    function filterDropdownList(searchTerm) {
      const items = dropdownList.querySelectorAll('.fitter-dropdown-item');
      items.forEach(item => {
        const name = (item.dataset.name || item.textContent).toLowerCase();
        if (name.includes(searchTerm) || searchTerm === '') {
          item.classList.remove('hidden');
        } else {
          item.classList.add('hidden');
        }
      });
    }
  }

  function populateFittersTable(fitters) {
    const fittersCount = document.getElementById('fittersCount');
    const fittersTableBody = document.getElementById('fittersTableBody');
    const fitterNameFilter = document.getElementById('fitterNameFilter');
    const showIncompleteOnly = document.getElementById('showIncompleteOnly');
    
    if (!fittersCount || !fittersTableBody) return;
    
    // Get filter values
    const selectedFitterId = fitterNameFilter ? fitterNameFilter.value : '';
    const showIncomplete = showIncompleteOnly ? showIncompleteOnly.checked : false;
    
    // Populate dropdown only once (when fitters are first loaded or list changes)
    const fittersHash = fitters.map(f => f.id).join(',');
    if (!window.fittersDropdownPopulated || window.fittersDropdownHash !== fittersHash) {
      populateFitterDropdown(fitters);
      window.fittersDropdownPopulated = true;
      window.fittersDropdownHash = fittersHash;
      
      // Restore selected value after populating
      if (selectedFitterId && fitterNameFilter) {
        fitterNameFilter.value = selectedFitterId;
        const fitterNameInput = document.getElementById('fitterNameFilterInput');
        const selectedFitter = fitters.find(f => f.id === selectedFitterId);
        if (selectedFitter && fitterNameInput) {
          fitterNameInput.value = selectedFitter.name;
        }
      }
    }
    
    // Update input display if a fitter is selected (only if not already set by user interaction)
    const fitterNameInput = document.getElementById('fitterNameFilterInput');
    if (fitterNameInput) {
      if (selectedFitterId) {
        const selectedFitter = fitters.find(f => f.id === selectedFitterId);
        if (selectedFitter && fitterNameInput.value !== selectedFitter.name) {
          fitterNameInput.value = selectedFitter.name;
        }
      } else if (!fitterNameInput.value || fitterNameInput.value === '') {
        fitterNameInput.value = 'All Fitters';
      }
    }
    
    // Sort fitters alphabetically by name
    let filteredFitters = [...fitters].sort((a, b) => {
      const nameA = a.name.toLowerCase();
      const nameB = b.name.toLowerCase();
      return nameA.localeCompare(nameB);
    });
    
    // Apply fitter name filter
    if (selectedFitterId) {
      filteredFitters = filteredFitters.filter(f => f.id === selectedFitterId);
    }
    
    // Apply incomplete filter
    if (showIncomplete) {
      filteredFitters = filteredFitters.filter(f => calculateAuditStatus(f) === 'incomplete');
    }
    
    fittersCount.textContent = `Showing ${filteredFitters.length} fitters`;
    
    if (filteredFitters.length === 0) {
      fittersTableBody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 2rem;">No fitters found</td></tr>';
      return;
    }
    
    fittersTableBody.innerHTML = filteredFitters.map(fitter => {
      // Helper function to get insurance status color class
      const getInsuranceStatusClass = (status) => {
        if (!status) return 'status-not-received';
        const statusLower = status.toLowerCase();
        if (statusLower.includes('completed') || statusLower.includes('valid')) return 'status-completed';
        return 'status-not-received';
      };

      // Helper function to get terms status color class
      const getTermsStatusClass = (status) => {
        if (!status) return 'status-not-received';
        const statusLower = status.toLowerCase();
        if (statusLower.includes('completed') || statusLower.includes('signed')) return 'status-completed';
        return 'status-not-received';
      };

      // Calculate audit status
      const auditStatus = calculateAuditStatus(fitter);
      const auditStatusText = auditStatus === 'complete' ? 'Complete' : 'Incomplete';

      return `
      <tr>
        <td>
          <input type="checkbox" class="fitter-checkbox" data-fitter-id="${fitter.id}" data-fitter-email="${fitter.email || ''}" data-fitter-name="${fitter.name}">
        </td>
        <td>
          <div class="action-buttons">
              <button class="btn-edit" onclick="openEditFitterModal('${fitter.id}', '${fitter.firstName} ${fitter.lastName}', '${JSON.stringify(Array.isArray(fitter.skillset) ? fitter.skillset : fitter.skillset.split(',').map(s => s.trim())).replace(/"/g, '&quot;')}', '${JSON.stringify(Array.isArray(fitter.skillsShortage) ? fitter.skillsShortage : fitter.skillsShortage.split(',').map(s => s.trim())).replace(/"/g, '&quot;')}', '${fitter.postcodesCovered || ''}', '${fitter.workConsistency || ''}')">Edit</button>
            <button class="btn-view" onclick="viewFitterInCRM('${fitter.id}')">View</button>
          </div>
        </td>
        <td>${fitter.name}</td>
        <td>${fitter.email}</td>
        <td>${fitter.phone}</td>
        <td>
          <span class="audit-status ${auditStatus}">${auditStatusText}</span>
        </td>
        <td>
          <div class="skillset-list">
            ${fitter.skillset && fitter.skillset !== '' ? 
              (Array.isArray(fitter.skillset) ? fitter.skillset : fitter.skillset.split(',')).map(skill => 
                `<span class="skill-tag data-available" title="${skill.trim()}">${skill.trim()}</span>`
              ).join('') : '<span class="skill-tag data-missing">Missing Data</span>'}
          </div>
        </td>
        <td><span class="${getDataStatusClass(fitter.postcodesCovered)}">${getDisplayValue(fitter.postcodesCovered)}</span></td>
        <td>
          <div class="skills-shortage-list">
            ${fitter.skillsShortage && fitter.skillsShortage !== '' ? 
              (Array.isArray(fitter.skillsShortage) ? fitter.skillsShortage : fitter.skillsShortage.split(',')).map(skill => 
                `<span class="skill-tag data-available" title="${skill.trim()}">${skill.trim()}</span>`
              ).join('') : '<span class="skill-tag data-missing">Missing Data</span>'}
          </div>
        </td>
        <td><span class="${getDataStatusClass(fitter.workConsistency)}">${getDisplayValue(fitter.workConsistency)}</span></td>
          <td>
            <span class="status-badge ${getTermsStatusClass(fitter.termsMasdouk)}">
              ${getDisplayValue(fitter.termsMasdouk)}
            </span>
          </td>
          <td>
            <span class="status-badge ${getInsuranceStatusClass(fitter.insuranceStatus)}">
              ${getDisplayValue(fitter.insuranceStatus)}
            </span>
          </td>
          <td>
            <div class="insurance-details">
              <div class="insurance-number ${getDataStatusClass(fitter.plInsuranceNo)}">${getDisplayValue(fitter.plInsuranceNo)}</div>
              <div class="pl-date ${getDataStatusClass(fitter.plDate)}">${fitter.plDate ? new Date(fitter.plDate).toLocaleDateString() : 'Missing Data'}</div>
            </div>
          </td>
      </tr>
      `;
    }).join('');
    
    // Set up checkbox event listeners after populating table
    setupFitterCheckboxes();
  }
  
  // =============================================
  // Fitter Selection and Email Functionality
  // =============================================
  
  function setupFitterCheckboxes() {
    // Select all checkbox - remove old listener and add new one
    const selectAllCheckbox = document.getElementById('selectAllFitters');
    if (selectAllCheckbox) {
      // Clone and replace to remove old event listeners
      const newSelectAll = selectAllCheckbox.cloneNode(true);
      selectAllCheckbox.parentNode.replaceChild(newSelectAll, selectAllCheckbox);
      
      newSelectAll.addEventListener('change', function() {
        const checkboxes = document.querySelectorAll('.fitter-checkbox');
        checkboxes.forEach(checkbox => {
          checkbox.checked = this.checked;
        });
        updateSendEmailButton();
      });
    }
    
    // Individual checkboxes - they're recreated each time, so just add listeners
    const checkboxes = document.querySelectorAll('.fitter-checkbox');
    checkboxes.forEach(checkbox => {
      checkbox.addEventListener('change', function() {
        updateSelectAllCheckbox();
        updateSendEmailButton();
      });
    });
    
    // Send Email button - ensure it only has one listener
    const sendEmailBtn = document.getElementById('sendEmailBtn');
    if (sendEmailBtn) {
      if (!sendEmailBtn.hasAttribute('data-email-listener-attached')) {
        sendEmailBtn.setAttribute('data-email-listener-attached', 'true');
        sendEmailBtn.addEventListener('click', function(e) {
          e.preventDefault();
          e.stopPropagation();
          console.log('Send Email button clicked');
          openSendEmailModal();
        });
        console.log('Send Email button listener attached');
      }
    } else {
      console.warn('Send Email button not found');
    }
    
    // Update button state initially
    updateSendEmailButton();
  }
  
  function updateSelectAllCheckbox() {
    const selectAllCheckbox = document.getElementById('selectAllFitters');
    const checkboxes = document.querySelectorAll('.fitter-checkbox');
    if (selectAllCheckbox && checkboxes.length > 0) {
      const allChecked = Array.from(checkboxes).every(cb => cb.checked);
      const someChecked = Array.from(checkboxes).some(cb => cb.checked);
      selectAllCheckbox.checked = allChecked;
      selectAllCheckbox.indeterminate = someChecked && !allChecked;
    }
  }
  
  function updateSendEmailButton() {
    const sendEmailBtn = document.getElementById('sendEmailBtn');
    const selectedCount = document.querySelectorAll('.fitter-checkbox:checked').length;
    if (sendEmailBtn) {
      sendEmailBtn.disabled = selectedCount === 0;
      sendEmailBtn.textContent = selectedCount > 0 ? `Send Email (${selectedCount})` : 'Send Email';
    }
  }
  
  function getSelectedFitters() {
    const checkboxes = document.querySelectorAll('.fitter-checkbox:checked');
    const selectedFitters = [];
    checkboxes.forEach(checkbox => {
      const email = checkbox.getAttribute('data-fitter-email');
      const name = checkbox.getAttribute('data-fitter-name');
      if (email && email.trim() !== '') {
        selectedFitters.push({
          id: checkbox.getAttribute('data-fitter-id'),
          email: email.trim(),
          name: name
        });
      }
    });
    return selectedFitters;
  }
  
  function openSendEmailModal() {
    console.log('openSendEmailModal called');
    const selectedFitters = getSelectedFitters();
    console.log('Selected fitters:', selectedFitters);
    
    if (selectedFitters.length === 0) {
      alert('Please select at least one fitter with an email address.');
      return;
    }
    
    const modal = document.getElementById('sendEmailModal');
    const recipientCount = document.getElementById('emailRecipientCount');
    
    if (!modal) {
      console.error('Send email modal not found!');
      alert('Email modal not found. Please refresh the page.');
      return;
    }
    
    if (!recipientCount) {
      console.error('Recipient count element not found!');
    }
    
    if (modal && recipientCount) {
      recipientCount.textContent = `${selectedFitters.length} ${selectedFitters.length === 1 ? 'fitter' : 'fitters'}`;
      modal.style.display = 'block';
      console.log('Modal displayed');
      
      // Clear form
      const subjectInput = document.getElementById('emailSubject');
      const messageInput = document.getElementById('emailMessage');
      if (subjectInput) subjectInput.value = '';
      if (messageInput) messageInput.value = '';
      updateEmailCharCount();
    } else {
      console.error('Modal or recipient count element not found');
    }
  }
  
  function closeSendEmailModal() {
    const modal = document.getElementById('sendEmailModal');
    if (modal) {
      modal.style.display = 'none';
    }
  }
  
  function updateEmailCharCount() {
    const message = document.getElementById('emailMessage').value;
    const charCount = document.getElementById('emailCharCount');
    if (charCount) {
      charCount.textContent = `${message.length} characters`;
    }
  }
  
  async function sendEmailToFitters(event) {
    console.log('sendEmailToFitters called', event);
    
    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }
    
    // Check if ZOHO SDK is available
    if (typeof ZOHO === 'undefined') {
      alert('Zoho CRM SDK is not available. Please ensure you are running this widget within Zoho CRM.');
      console.error('ZOHO object not defined');
      return;
    }
    
    if (!ZOHO.CRM || !ZOHO.CRM.FUNCTIONS) {
      alert('Zoho CRM API is not available. Please ensure you are running this widget within Zoho CRM.');
      console.error('ZOHO.CRM.FUNCTIONS not available', ZOHO);
      return;
    }
    
    const selectedFitters = getSelectedFitters();
    if (selectedFitters.length === 0) {
      alert('Please select at least one fitter with an email address.');
      return;
    }
    
    const subject = document.getElementById('emailSubject').value.trim();
    const message = document.getElementById('emailMessage').value.trim();
    
    if (!subject || !message) {
      alert('Please fill in both subject and message.');
      return;
    }
    
    // Convert line breaks to <br> tags for HTML email
    const messageWithBreaks = message.replace(/\n/g, '<br>');
    
    // Get emails as comma-separated string
    const emails = selectedFitters.map(f => f.email).join(',');
    
    console.log('📧 Preparing to send email...');
    console.log('Subject:', subject);
    console.log('Message (original):', message);
    console.log('Recipients:', emails);
    console.log('Number of recipients:', selectedFitters.length);
    console.log('Message (with breaks):', messageWithBreaks);
    
    const sendBtn = document.getElementById('submitSendEmail');
    if (!sendBtn) {
      console.error('Send button not found');
      return;
    }
    
    const originalText = sendBtn.innerHTML;
    sendBtn.disabled = true;
    sendBtn.innerHTML = 'Sending...';
    
    try {
      // Call the Zoho CRM function
      const funcName = 'sendMailToFitters';
      const reqData = {
        "arguments": JSON.stringify({
          "fitter_ids": emails, // Passing emails instead of IDs
          "subject_str": subject,
          "message_str": messageWithBreaks
        })
      };
      
      console.log('🚀 Calling function:', funcName);
      console.log('📦 Request data:', reqData);
      
      const response = await ZOHO.CRM.FUNCTIONS.execute(funcName, reqData);
      
      console.log('✅ Function response:', response);
      
      if (response && response.code === 'success') {
        alert(`Emails sent successfully to ${selectedFitters.length} ${selectedFitters.length === 1 ? 'fitter' : 'fitters'}!`);
        
        // Clear the form and close modal
        document.getElementById('emailSubject').value = '';
        document.getElementById('emailMessage').value = '';
        updateEmailCharCount();
        closeSendEmailModal();
        
        // Uncheck all checkboxes
        document.querySelectorAll('.fitter-checkbox').forEach(cb => cb.checked = false);
        const selectAllCheckbox = document.getElementById('selectAllFitters');
        if (selectAllCheckbox) {
          selectAllCheckbox.checked = false;
          selectAllCheckbox.indeterminate = false;
        }
        updateSendEmailButton();
      } else {
        const errorMsg = response && response.message ? response.message : 'Unknown error';
        console.error('Failed to send emails:', errorMsg, response);
        alert('Failed to send emails: ' + errorMsg);
      }
      
    } catch (error) {
      console.error('❌ Error calling function:', error);
      const errorMessage = error.message || error.toString() || 'Unknown error occurred';
      alert('Error sending emails: ' + errorMessage);
    } finally {
      sendBtn.disabled = false;
      sendBtn.innerHTML = originalText;
    }
  }
  
  // =============================================
  // Edit Fitter Modal Functionality
  // =============================================
  
  window.openEditFitterModal = function(fitterId, fitterName, skillsetJson, skillsShortageJson, postcodeArea, workConsistency) {
    const modal = document.getElementById('editFitterModal');
    const nameInput = document.getElementById('editFitterName');
    const skillsetSelected = document.getElementById('editSkillsetSelected');
    const skillsShortageSelected = document.getElementById('editSkillsShortageSelected');
    const postcodeInput = document.getElementById('editFitterPostcode');
    const workConsistencySelect = document.getElementById('editFitterWorkConsistency');
    
    if (!modal || !nameInput || !skillsetSelected || !skillsShortageSelected || !postcodeInput || !workConsistencySelect) {
      console.error('Edit modal elements not found!');
      return;
    }
    
    // Set fitter name
    nameInput.value = fitterName;
    
    // Parse skillset and skills shortage
    const currentSkillset = JSON.parse(skillsetJson.replace(/&quot;/g, '"'));
    const currentSkillsShortage = JSON.parse(skillsShortageJson.replace(/&quot;/g, '"'));
    
    // Set postcode and work consistency
    postcodeInput.value = postcodeArea || '';
    workConsistencySelect.value = workConsistency || '';
    
    // Store current fitter data
    window.currentEditingFitter = {
      id: fitterId,
      name: fitterName,
      skillset: currentSkillset,
      skillsShortage: currentSkillsShortage,
      postcodeArea: postcodeArea || '',
      workConsistency: workConsistency || ''
    };
    
    // Populate multiselects
    populateEditSkillsetMultiselect(currentSkillset);
    populateEditSkillsShortageMultiselect(currentSkillsShortage);
    
    // Populate work consistency dropdown with unique values
    populateWorkConsistencyDropdown();
    
    // Show modal
    modal.style.display = 'block';
  }
  
  function populateEditSkillsetMultiselect(selectedSkillset) {
    const selectedDiv = document.getElementById('editSkillsetSelected');
    const optionsDiv = document.getElementById('editSkillsetOptions');
    
    if (!selectedDiv || !optionsDiv) return;
    
    // Use the same logic as calendar filters to get unique skillsets
    const allSkillsets = [...new Set(window.fitters.flatMap(f => {
      const skillsetField = f.skillset;
      
      if (!skillsetField) {
        return [];
      }
      
      let skills = [];
      if (Array.isArray(skillsetField)) {
        // If it's already an array, use it directly
        skills = skillsetField.map(skill => skill.trim()).filter(skill => skill.length > 0);
      } else if (typeof skillsetField === 'string') {
        // If it's a string, split it
        skills = skillsetField.split(/[;,]/).map(skill => skill.trim()).filter(skill => skill.length > 0);
      }
      
      return skills;
    }))];
    
    // Sort alphabetically for consistency
    allSkillsets.sort((a, b) => a.localeCompare(b));
    
    // Clear existing content
    selectedDiv.innerHTML = '';
    optionsDiv.innerHTML = '';
    
    // Add selected skillsets as tags
    selectedSkillset.forEach(skill => {
      const tag = document.createElement('span');
      tag.className = 'multiselect-tag';
      tag.innerHTML = `${skill} <span class="remove" onclick="removeSkillsetTag(this)">&times;</span>`;
      selectedDiv.appendChild(tag);
    });
    
    if (selectedSkillset.length === 0) {
      selectedDiv.innerHTML = '<span style="color: #9ca3af;">Select skillsets...</span>';
    }
    
    // Add options (filter out already selected ones)
    allSkillsets.forEach(skill => {
      // Skip if this skill is already selected
      if (selectedSkillset.includes(skill)) {
        return;
      }
      
      const option = document.createElement('div');
      option.className = 'multiselect-option';
      option.textContent = skill;
      option.onclick = () => toggleSkillsetOption(option, skill);
      optionsDiv.appendChild(option);
    });
  }
  
  function populateEditSkillsShortageMultiselect(selectedSkillsShortage) {
    const selectedDiv = document.getElementById('editSkillsShortageSelected');
    const optionsDiv = document.getElementById('editSkillsShortageOptions');
    
    if (!selectedDiv || !optionsDiv) return;
    
    // Use the same logic as calendar filters to get unique skills shortages
    const allSkillsShortages = [...new Set(window.fitters.flatMap(f => {
      const skillsShortageField = f.skillsShortage;
      
      if (!skillsShortageField || skillsShortageField === '') {
        return [];
      }
      
      let shortages = [];
      if (Array.isArray(skillsShortageField)) {
        // If it's already an array, process each element
        shortages = skillsShortageField.flatMap(item => {
          if (typeof item === 'string') {
            return parseSkillsShortageString(item);
          }
          return [item];
        }).filter(shortage => shortage.length > 0);
      } else if (typeof skillsShortageField === 'string') {
        // If it's a string, split it
        shortages = skillsShortageField.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
      } else if (typeof skillsShortageField === 'object' && skillsShortageField !== null) {
        // If it's an object, try to extract the value
        if (skillsShortageField.value) {
          shortages = skillsShortageField.value.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
        } else if (Object.keys(skillsShortageField).length === 1) {
          const value = Object.values(skillsShortageField)[0];
          if (typeof value === 'string') {
            shortages = value.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
          }
        } else {
          const stringValues = Object.values(skillsShortageField).filter(v => typeof v === 'string' && v.trim() !== '');
          if (stringValues.length > 0) {
            shortages = stringValues.flatMap(v => v.split(/[;,]/).map(shortage => shortage.trim())).filter(shortage => shortage.length > 0);
          }
        }
      }
      
      return shortages;
    }))];
    
    // Sort alphabetically for consistency
    allSkillsShortages.sort((a, b) => a.localeCompare(b));
    
    // Clear existing content
    selectedDiv.innerHTML = '';
    optionsDiv.innerHTML = '';
    
    // Add selected skills shortages as tags
    selectedSkillsShortage.forEach(skill => {
      const tag = document.createElement('span');
      tag.className = 'multiselect-tag';
      tag.innerHTML = `${skill} <span class="remove" onclick="removeSkillsShortageTag(this)">&times;</span>`;
      selectedDiv.appendChild(tag);
    });
    
    if (selectedSkillsShortage.length === 0) {
      selectedDiv.innerHTML = '<span style="color: #9ca3af;">Select skills shortages...</span>';
    }
    
    // Add options (filter out already selected ones)
    allSkillsShortages.forEach(skill => {
      // Skip if this skill shortage is already selected
      if (selectedSkillsShortage.includes(skill)) {
        return;
      }
      
      const option = document.createElement('div');
      option.className = 'multiselect-option';
      option.textContent = skill;
      option.onclick = () => toggleSkillsShortageOption(option, skill);
      optionsDiv.appendChild(option);
    });
  }
  
  function toggleSkillsetOption(option, skill) {
    const selectedDiv = document.getElementById('editSkillsetSelected');
    const isSelected = option.classList.contains('selected');
    
    if (isSelected) {
      option.classList.remove('selected');
      // Remove tag
      const tags = selectedDiv.querySelectorAll('.multiselect-tag');
      tags.forEach(tag => {
        if (tag.textContent.trim().startsWith(skill)) {
          tag.remove();
        }
      });
    } else {
      option.classList.add('selected');
      // Add tag
      const tag = document.createElement('span');
      tag.className = 'multiselect-tag';
      tag.innerHTML = `${skill} <span class="remove" onclick="removeSkillsetTag(this)">&times;</span>`;
      selectedDiv.appendChild(tag);
    }
    
    // Update placeholder
    if (selectedDiv.querySelectorAll('.multiselect-tag').length === 0) {
      selectedDiv.innerHTML = '<span style="color: #9ca3af;">Select skillsets...</span>';
    }
    
    // Refresh options list to show/hide selected items
    refreshSkillsetOptions();
  }
  
  function toggleSkillsShortageOption(option, skill) {
    const selectedDiv = document.getElementById('editSkillsShortageSelected');
    const isSelected = option.classList.contains('selected');
    
    if (isSelected) {
      option.classList.remove('selected');
      // Remove tag
      const tags = selectedDiv.querySelectorAll('.multiselect-tag');
      tags.forEach(tag => {
        if (tag.textContent.trim().startsWith(skill)) {
          tag.remove();
        }
      });
    } else {
      option.classList.add('selected');
      // Add tag
      const tag = document.createElement('span');
      tag.className = 'multiselect-tag';
      tag.innerHTML = `${skill} <span class="remove" onclick="removeSkillsShortageTag(this)">&times;</span>`;
      selectedDiv.appendChild(tag);
    }
    
    // Update placeholder
    if (selectedDiv.querySelectorAll('.multiselect-tag').length === 0) {
      selectedDiv.innerHTML = '<span style="color: #9ca3af;">Select skills shortages...</span>';
    }
    
    // Refresh options list to show/hide selected items
    refreshSkillsShortageOptions();
  }
  
  // Helper function to refresh skillset options
  function refreshSkillsetOptions() {
    const selectedDiv = document.getElementById('editSkillsetSelected');
    const optionsDiv = document.getElementById('editSkillsetOptions');
    
    if (!selectedDiv || !optionsDiv) return;
    
    // Get currently selected skillsets
    const selectedSkillsets = Array.from(selectedDiv.querySelectorAll('.multiselect-tag'))
      .map(tag => tag.textContent.replace('×', '').trim());
    
    // Get all available skillsets
    const allSkillsets = [...new Set(window.fitters.flatMap(f => {
      const skillsetField = f.skillset;
      if (!skillsetField) return [];
      
      let skills = [];
      if (Array.isArray(skillsetField)) {
        skills = skillsetField.map(skill => skill.trim()).filter(skill => skill.length > 0);
      } else if (typeof skillsetField === 'string') {
        skills = skillsetField.split(/[;,]/).map(skill => skill.trim()).filter(skill => skill.length > 0);
      }
      return skills;
    }))];
    
    // Clear and repopulate options
    optionsDiv.innerHTML = '';
    allSkillsets.forEach(skill => {
      // Skip if this skill is already selected
      if (selectedSkillsets.includes(skill)) {
        return;
      }
      
      const option = document.createElement('div');
      option.className = 'multiselect-option';
      option.textContent = skill;
      option.onclick = () => toggleSkillsetOption(option, skill);
      optionsDiv.appendChild(option);
    });
  }
  
  // Helper function to refresh skills shortage options
  function refreshSkillsShortageOptions() {
    const selectedDiv = document.getElementById('editSkillsShortageSelected');
    const optionsDiv = document.getElementById('editSkillsShortageOptions');
    
    if (!selectedDiv || !optionsDiv) return;
    
    // Get currently selected skills shortages
    const selectedSkillsShortages = Array.from(selectedDiv.querySelectorAll('.multiselect-tag'))
      .map(tag => tag.textContent.replace('×', '').trim());
    
    // Get all available skills shortages using the same logic as calendar filters
    const allSkillsShortages = [...new Set(window.fitters.flatMap(f => {
      const skillsShortageField = f.skillsShortage;
      
      if (!skillsShortageField || skillsShortageField === '') {
        return [];
      }
      
      let shortages = [];
      if (Array.isArray(skillsShortageField)) {
        shortages = skillsShortageField.flatMap(item => {
          if (typeof item === 'string') {
            return parseSkillsShortageString(item);
          }
          return [item];
        }).filter(shortage => shortage.length > 0);
      } else if (typeof skillsShortageField === 'string') {
        shortages = skillsShortageField.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
      } else if (typeof skillsShortageField === 'object' && skillsShortageField !== null) {
        if (skillsShortageField.value) {
          shortages = skillsShortageField.value.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
        } else if (Object.keys(skillsShortageField).length === 1) {
          const value = Object.values(skillsShortageField)[0];
          if (typeof value === 'string') {
            shortages = value.split(/[;,]/).map(shortage => shortage.trim()).filter(shortage => shortage.length > 0);
          }
        } else {
          const stringValues = Object.values(skillsShortageField).filter(v => typeof v === 'string' && v.trim() !== '');
          if (stringValues.length > 0) {
            shortages = stringValues.flatMap(v => v.split(/[;,]/).map(shortage => shortage.trim())).filter(shortage => shortage.length > 0);
          }
        }
      }
      
      return shortages;
    }))];
    
    // Sort alphabetically for consistency
    allSkillsShortages.sort((a, b) => a.localeCompare(b));
    
    // Clear and repopulate options
    optionsDiv.innerHTML = '';
    allSkillsShortages.forEach(skill => {
      // Skip if this skill shortage is already selected
      if (selectedSkillsShortages.includes(skill)) {
        return;
      }
      
      const option = document.createElement('div');
      option.className = 'multiselect-option';
      option.textContent = skill;
      option.onclick = () => toggleSkillsShortageOption(option, skill);
      optionsDiv.appendChild(option);
    });
  }
  
  window.removeSkillsetTag = function(tagElement) {
    const skill = tagElement.parentElement.textContent.replace('×', '').trim();
    tagElement.parentElement.remove();
    
    // Update placeholder
    const selectedDiv = document.getElementById('editSkillsetSelected');
    if (selectedDiv.querySelectorAll('.multiselect-tag').length === 0) {
      selectedDiv.innerHTML = '<span style="color: #9ca3af;">Select skillsets...</span>';
    }
    
    // Refresh options list to show the removed item
    refreshSkillsetOptions();
  }
  
  window.removeSkillsShortageTag = function(tagElement) {
    const skill = tagElement.parentElement.textContent.replace('×', '').trim();
    tagElement.parentElement.remove();
    
    // Update placeholder
    const selectedDiv = document.getElementById('editSkillsShortageSelected');
    if (selectedDiv.querySelectorAll('.multiselect-tag').length === 0) {
      selectedDiv.innerHTML = '<span style="color: #9ca3af;">Select skills shortages...</span>';
    }
    
    // Refresh options list to show the removed item
    refreshSkillsShortageOptions();
  }
  
  // Function to populate work consistency dropdown with unique values
  function populateWorkConsistencyDropdown() {
    const workConsistencySelect = document.getElementById('editFitterWorkConsistency');
    if (!workConsistencySelect) return;
    
    // Get current value to preserve selection
    const currentValue = workConsistencySelect.value;
    
    // Clear existing options except the first one
    workConsistencySelect.innerHTML = '<option value="">Select work consistency...</option>';
    
    // Get all unique work consistencies from fitters data
    const allWorkConsistencies = [...new Set(window.fitters
      .map(f => f.workConsistency)
      .filter(wc => wc && wc.trim() !== '')
      .sort()
    )];
    
    console.log('Unique work consistencies found:', allWorkConsistencies);
    
    // Add unique options
    allWorkConsistencies.forEach(consistency => {
      const option = document.createElement('option');
      option.value = consistency;
      option.textContent = consistency;
      workConsistencySelect.appendChild(option);
    });
    
    // Restore the current value if it exists
    if (currentValue) {
      workConsistencySelect.value = currentValue;
    }
  }
  
  window.viewFitterInCRM = function(fitterId) {
    // Open fitter in Zoho CRM in new tab
    try {
      const context = window.ZOHO.embeddedApp.getContext();
      const orgId = context ? context.orgId : 'default';
      const crmUrl = `https://crm.zoho.com/crm/org${orgId}/tab/Contacts/${fitterId}`;
      window.open(crmUrl, '_blank');
    } catch (error) {
      console.error('Error getting context, using default URL:', error);
      // Fallback URL without orgId
      const crmUrl = `https://crm.zoho.com/crm/tab/Contacts/${fitterId}`;
      window.open(crmUrl, '_blank');
    }
  }
  
  // =============================================
  // Modal Event Listeners
  // =============================================
  
  function initializeEditModal() {
    const modal = document.getElementById('editFitterModal');
    const closeBtn = document.getElementById('closeEditModal');
    const cancelBtn = document.getElementById('cancelEditFitter');
    const saveBtn = document.getElementById('saveEditFitter');
    const form = document.getElementById('editFitterForm');
    
    if (!modal || !closeBtn || !cancelBtn || !saveBtn || !form) {
      console.error('Edit modal elements not found!');
      return;
    }
    
    // Close modal events
    closeBtn.onclick = () => modal.style.display = 'none';
    cancelBtn.onclick = () => modal.style.display = 'none';
    
    // Click outside modal to close
    modal.onclick = (e) => {
      if (e.target === modal) {
        modal.style.display = 'none';
      }
    };
    
    // Form submission
    form.onsubmit = (e) => {
      e.preventDefault();
      saveFitterChanges();
    };
    
    // Multiselect dropdown toggles
    const skillsetDropdown = document.getElementById('editSkillsetDropdown');
    const skillsShortageDropdown = document.getElementById('editSkillsShortageDropdown');
    
    if (skillsetDropdown) {
      skillsetDropdown.onclick = (e) => {
        e.stopPropagation();
        const options = document.getElementById('editSkillsetOptions');
        options.classList.toggle('show');
      };
    }
    
    if (skillsShortageDropdown) {
      skillsShortageDropdown.onclick = (e) => {
        e.stopPropagation();
        const options = document.getElementById('editSkillsShortageOptions');
        options.classList.toggle('show');
      };
    }
    
    // Close dropdowns when clicking outside
    document.addEventListener('click', () => {
      const skillsetOptions = document.getElementById('editSkillsetOptions');
      const skillsShortageOptions = document.getElementById('editSkillsShortageOptions');
      if (skillsetOptions) skillsetOptions.classList.remove('show');
      if (skillsShortageOptions) skillsShortageOptions.classList.remove('show');
    });
  }
  
  async function saveFitterChanges() {
    if (!window.currentEditingFitter) {
      console.error('No fitter being edited');
      return;
    }
    
    const selectedDiv = document.getElementById('editSkillsetSelected');
    const skillsShortageDiv = document.getElementById('editSkillsShortageSelected');
    const postcodeInput = document.getElementById('editFitterPostcode');
    const workConsistencySelect = document.getElementById('editFitterWorkConsistency');
    
    if (!selectedDiv || !skillsShortageDiv || !postcodeInput || !workConsistencySelect) {
      console.error('Edit modal elements not found');
      return;
    }
    
    // Get selected skillsets
    const skillsetTags = selectedDiv.querySelectorAll('.multiselect-tag');
    const newSkillset = Array.from(skillsetTags).map(tag => 
      tag.textContent.replace('×', '').trim()
    );
    
    // Get selected skills shortages
    const skillsShortageTags = skillsShortageDiv.querySelectorAll('.multiselect-tag');
    const newSkillsShortage = Array.from(skillsShortageTags).map(tag => 
      tag.textContent.replace('×', '').trim()
    );
    
    // Get postcode and work consistency
    const newPostcodeArea = postcodeInput.value.trim();
    const newWorkConsistency = workConsistencySelect.value;
    
    try {
      console.log('Updating fitter:', window.currentEditingFitter.id);
      console.log('New skillset:', newSkillset);
      console.log('New skills shortage:', newSkillsShortage);
      console.log('New postcode area:', newPostcodeArea);
      console.log('New work consistency:', newWorkConsistency);
      
      // Update fitter in CRM
      const updateData = {
        id: window.currentEditingFitter.id,
        Fitter_Skillset_2: newSkillset,
        Skills_They_Can_t_Do: newSkillsShortage,
        Postcode_Area_New: newPostcodeArea,
        Work_Consistency: newWorkConsistency
      };
      
      console.log('Update data:', updateData);
      
      const response = await ZOHO.CRM.API.updateRecord({
        Entity: "Contacts",
        APIData: updateData
      });
      
      console.log('Update response:', response);
      
      if (response && response.data && response.data.length > 0) {
        // Update local data
        const fitterIndex = window.fitters.findIndex(f => f.id === window.currentEditingFitter.id);
        if (fitterIndex !== -1) {
          window.fitters[fitterIndex].skillset = newSkillset;
          window.fitters[fitterIndex].skillsShortage = newSkillsShortage;
          window.fitters[fitterIndex].postcodesCovered = newPostcodeArea;
          window.fitters[fitterIndex].workConsistency = newWorkConsistency;
        }
        
        // Refresh fitters table
        populateFittersTable(window.fitters);
        
        // Close modal
        document.getElementById('editFitterModal').style.display = 'none';
        
        // Show success message
        alert('Fitter updated successfully!');
        console.log('Fitter update completed successfully');
      } else {
        console.error('Update failed - no data in response:', response);
        throw new Error('Update failed - no data in response');
      }
    } catch (error) {
      console.error('Error updating fitter:', error);
      console.error('Error details:', error.message);
      alert(`Error updating fitter: ${error.message}. Please try again.`);
    }
  }
  
  // =============================================
  // Public API for Zoho Integration
  // =============================================
  
  // Function to be called from Zoho CRM
  window.updateCalendarData = function(newFitters, newOrders) {
    if (newFitters) {
      fitters = newFitters;
      populateFilterOptions();
    }
    
    if (newOrders) {
      orders = newOrders;
    }
    
    updateCalendar();
  };
  
  // Function to refresh calendar
  window.refreshCalendar = function() {
    loadCalendarData();
  };
  
  console.log('Fitters Dashboard Calendar Widget: Loaded successfully');
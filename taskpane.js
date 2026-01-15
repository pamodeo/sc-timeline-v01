/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById('btnSync').onclick = syncToTimeline;
    document.getElementById('btnCancel').onclick = closePane;
    
    const activityType = document.getElementById('activityType');
    const engagementType = document.getElementById('engagementType');
    const customerEvent = document.getElementById('customerEvent');
    
    activityType.addEventListener('change', validateForm);
    engagementType.addEventListener('change', validateForm);
    customerEvent.addEventListener('input', validateForm);
    
    // Load existing values
    loadExistingValues();
  }
});

function validateForm() {
  const activityType = document.getElementById('activityType').value;
  const engagementType = document.getElementById('engagementType').value;
  const customerEvent = document.getElementById('customerEvent').value;
  const btnSync = document.getElementById('btnSync');
  
  // Disable engagement type if PTO is selected
  if (activityType === 'PTO') {
    document.getElementById('engagementType').value = '';
    document.getElementById('engagementType').disabled = true;
  } else {
    document.getElementById('engagementType').disabled = false;
  }
  
  // Enable sync button logic
  let isValid = false;
  if (activityType && customerEvent) {
    if (activityType === 'PTO') {
      isValid = true;
    } else if (engagementType) {
      isValid = true;
    }
  }
  
  btnSync.disabled = !isValid;
}

function loadExistingValues() {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const customProps = result.value;
      
      const activityType = customProps.get('ActivityType');
      const engagementType = customProps.get('EngagementType');
      const customerEvent = customProps.get('CustomerEvent');
      const onSite = customProps.get('OnSite');
      const custInteraction = customProps.get('CustInteraction');
      const clevel = customProps.get('Clevel');
      
      if (activityType) document.getElementById('activityType').value = activityType;
      if (engagementType) document.getElementById('engagementType').value = engagementType;
      if (customerEvent) document.getElementById('customerEvent').value = customerEvent;
      if (onSite === true || onSite === 'true') document.getElementById('onSite').checked = true;
      if (custInteraction === true || custInteraction === 'true') document.getElementById('custInteraction').checked = true;
      if (clevel === true || clevel === 'true') document.getElementById('clevel').checked = true;
      
      validateForm();
    }
  });
}

function saveCustomProperties(callback) {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const customProps = result.value;
      
      customProps.set('ActivityType', document.getElementById('activityType').value);
      customProps.set('EngagementType', document.getElementById('engagementType').value);
      customProps.set('CustomerEvent', document.getElementById('customerEvent').value);
      customProps.set('OnSite', document.getElementById('onSite').checked);
      customProps.set('CustInteraction', document.getElementById('custInteraction').checked);
      customProps.set('Clevel', document.getElementById('clevel').checked);
      
      customProps.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          callback(true);
        } else {
          callback(false);
        }
      });
    }
  });
}

async function syncToTimeline() {
  const statusDiv = document.getElementById('status');
  statusDiv.className = 'status-message';
  statusDiv.style.display = 'none';
  
  // Save properties first
  saveCustomProperties(async (success) => {
    if (!success) {
      showStatus('Failed to save properties', 'error');
      return;
    }
    
    try {
      const item = Office.context.mailbox.item;
      
      // Get appointment data
      const appointmentData = await getAppointmentData(item);
      
      // Build JSON payload
      const json = buildJsonPayload(appointmentData);
      
      // Send to API
      const response = await fetch('https://dataflow-inbound-message-prd-euc1.eam.hxgnsmartcloud.com/api/message?tag=timeline', {
        method: 'POST',
        headers: {
          'accept': 'application/json',
          'X-Tenant-Id': 'HXGNDEMO0016_DEM',
          'Authorization': 'Basic SDNBV0JNX0hYR05ERU1PMDAxNl9ERU06RyFvYmEhMjAyMA==',
          'Content-Type': 'text/plain'
        },
        body: json
      });
      
      if (response.ok) {
        showStatus('Appointment sent to Timeline successfully!', 'success');
        const openTenant = confirm('Appointment sent to Databridge Pro for Timeline process sync.\\n\\nSelect OK to login to the Timeline tenant\nSelect Cancel to close');
        if (openTenant) {
          window.open('https://eu1.eam.hxgnsmartcloud.com/web/base/logindisp?tenant=HXGNDEMO0016_DEM', '_blank');
        }
      } else {
        const errorText = await response.text();
		const msgTextErr1 = 'Error: ' + ${response.status} + ' - ' ${errorText};
        showStatus(msgTextErr1, 'error');
      }
    } catch (error) {
		const msgTextErr2 = 'Error: ' + ${error.message};
      showStatus(msgTextErr2, 'error');
    }
  });
}

async function getAppointmentData(item) {
  return new Promise((resolve) => {
    const data = {
      subject: item.subject,
      location: item.location,
      start: item.start,
      end: item.end,
      organizer: '',
      body: ''
    };
    
    // Get organizer
    if (item.organizer) {
      data.organizer = item.organizer.emailAddress || item.organizer.displayName || '';
    }
    
    // Get body
    item.body.getAsync(Office.CoerceType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        data.body = result.value || '';
      }
      
      // Get custom properties
      item.loadCustomPropertiesAsync((propResult) => {
        if (propResult.status === Office.AsyncResultStatus.Succeeded) {
          const customProps = propResult.value;
          data.activityType = customProps.get('ActivityType') || '';
          data.engagementType = customProps.get('EngagementType') || '';
          data.customerEvent = customProps.get('CustomerEvent') || data.subject;
          data.onSite = customProps.get('OnSite') || false;
          data.custInteraction = customProps.get('CustInteraction') || false;
          data.clevel = customProps.get('Clevel') || false;
        }
        resolve(data);
      });
    });
  });
}

function buildJsonPayload(data) {
  // Get owner email (use current user as fallback)
  const ownerEmail = Office.context.mailbox.userProfile.emailAddress;
  
  // Parse organizer info
  let aliasStr = 'no alias';
  let firstNameStr = 'External';
  let lastNameStr = 'External';
  
  if (data.organizer) {
    if (data.organizer.includes('@')) {
      aliasStr = data.organizer;
      lastNameStr = data.organizer.split('@')[0];
      firstNameStr = '';
    } else {
      const parts = data.organizer.split(' ');
      if (parts.length >= 2) {
        firstNameStr = parts[0];
        lastNameStr = parts.slice(1).join(' ');
      } else {
        lastNameStr = data.organizer;
      }
    }
  }
  
  // Format dates
  const formatDate = (date) => {
    const d = new Date(date);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    const seconds = String(d.getSeconds()).padStart(2, '0');
	const myDate = ${day}+'/'+${month}+'/' + ${year}+' '+ ${hours}+':'+${minutes}+':'+${seconds};
    return myDate;
  };
  
  // Clean body
  let cleanBody = data.body.replace(/\\r\\n/g, ' ').replace(/\\n/g, ' ');
  cleanBody = cleanBody.replace(/"/g, ' ').replace(/[{}\\[\\]]/g, ' ');
  cleanBody = cleanBody.substring(0, 255);
  
  // Handle PTO
  let customerEvent = data.customerEvent;
  let engagementType = data.engagementType;
  if (data.activityType === 'PTO') {
    customerEvent = 'Personal Time OFF';
    engagementType = '';
  }
  
  const payload = {
    EntryID: Office.context.mailbox.item.itemId || '',
    globalID: Office.context.mailbox.item.itemId || '',
    Organizer: data.organizer,
    AuthorAlias: aliasStr,
    AuthorFirstname: firstNameStr,
    AuthorLastname: lastNameStr,
    OwnerEmail: ownerEmail,
    Subject: customerEvent,
    Start: formatDate(data.start),
    End: formatDate(data.end),
    Location: data.location || '',
    CreationTime: new Date().toISOString(),
    ActivityType: data.activityType,
    EngagementType: engagementType,
    OnSite: data.onSite.toString(),
    CustInteraction: data.custInteraction.toString(),
    Clevel: data.clevel.toString(),
    Note: cleanBody
  };
  
  return JSON.stringify(payload);
}

function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = 'status-message '+ ${type};
  statusDiv.style.display = 'block';
}

function closePane() {
  Office.context.ui.closeContainer();

}
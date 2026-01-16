/* global Office  v.25 */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById('btnSync').onclick = syncToTimeline;
    document.getElementById('btnCancel').onclick = closePane;
    
    const activityType = document.getElementById('activityType');
    const engagementType = document.getElementById('engagementType');
    const customerEvent = document.getElementById('customerEvent');
    
    activityType.addEventListener('change', validateForm);
    engagementType.addEventListener('change',  validateForm);
    customerEvent.addEventListener('input', validateForm);
    
	// Initial load
    loadExistingValues;
    
    // Run validation
    validateForm;
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
      const OnSite = customProps.get('OnSite');
      const CustInteraction = customProps.get('CustInteraction');
      const Clevel = customProps.get('Clevel');
      
      if (activityType) document.getElementById('activityType').value = activityType;
      if (engagementType) document.getElementById('engagementType').value = engagementType;
      if (customerEvent) document.getElementById('customerEvent').value = customerEvent;
      if (OnSite === true || OnSite === 'true') document.getElementById('OnSite').checked = true;
      if (CustInteraction === true || CustInteraction === 'true') document.getElementById('CustInteraction').checked = true;
      if (Clevel === true || Clevel === 'true') document.getElementById('Clevel').checked = true;
      
      // CRITICAL: Call validation AFTER the values are set
      validateForm;
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
      customProps.set('OnSite', document.getElementById('OnSite').checked);
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
      console.log("data.organizer:", appointmentData.organizer);
      
      // Build JSON payload
      const json = await buildJsonPayload(appointmentData);
      
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
        // showStatus('Appointment sent to Timeline successfully!\nClick on "Open Timeline Tenant" or "Close"', 'success');
        const msgText = await response.text();
		showStatus(msgText, 'success'); 
      } else {
        const errorText = await response.text();
		const msgTextErr1 = `Error: ${response.status} - ${errorText}`;
        showStatus(msgTextErr1, 'error');
      }
    } catch (error) {
		const msgTextErr2 = `Error: ${error.message}`;
        showStatus(msgTextErr2, 'error');
    }
  });
}

async function getAppointmentData(item) {
  return new Promise(async (resolve) => {
    // Helper to handle the "Compose vs Read" difference for strings/dates
    const getValue = async (property) => {
      if (property && typeof property.getAsync === 'function') {
        return new Promise(r => property.getAsync(result => r(result.value || '')));
      }
      return property || '';
    };
	  
    const data = {
      subject: await getValue(item.subject),
      location: await getValue(item.location),
      start: await getValue(item.start),
      end: await getValue(item.end),
      organizer: '',
      body: ''
    };

    // 1. Handle Organizer (It's a bit more complex)
    if (item.organizer) {
      if (typeof item.organizer.getAsync === 'function') {
        const orgRes = await new Promise(r => item.organizer.getAsync(r));
        data.organizer = orgRes.value ? (orgRes.value.emailAddress || orgRes.value.displayName) : '';
      } else {
        data.organizer = item.organizer.emailAddress || item.organizer.displayName || '';
      }
    }

    // 2. Get Body (Always Async)
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        data.body = result.value || '';
      }

      // 3. Get Custom Properties
      item.loadCustomPropertiesAsync((propResult) => {
        if (propResult.status === Office.AsyncResultStatus.Succeeded) {
          const props = propResult.value;
          // Keeping keys short (8-char logic)
          data.actType = props.get('ActivityType') || '';
          data.engType = props.get('EngagementType') || '';
          data.custEvt = props.get('CustomerEvent') || data.subject;
		  data.OnSite = props.get('OnSite');
		  data.CustInteraction = props.get('CustInteraction');
		  data.Clevel = props.get('Clevel');
        }
        resolve(data);
      });
    });
  });
}

async function buildJsonPayload(data) {
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
	const myDate = `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
    return myDate;
  };
  
  // Clean body
  let cleanBody = data.body.replace(/[\r\n]+/g, ' ').replace(/"/g, ' ').replace(/[{}\[\]]/g, ' ').substring(0, 255);  
// Handle PTO
  let customerEvent = data.custEvt; // Start with the captured subject
  let engagementType = data.engType;
  if (data.actType === 'PTO') {
    customerEvent = 'Personal Time OFF'; // Use the local variable
    engagementType = '';
  }
  const Location =  data.location || '';
  const CreationTime = new Date().toISOString();
  const OnSite = (data.OnSite ?? false).toString();
  const CustInteraction = (data.CustInteraction ?? false).toString();
  const Clevel = (data.Clevel ?? false).toString();

// EntryID (Standard Office.js itemId)
// Logic fix for IDs:
  const entryID = Office.context.mailbox.item.itemId || await new Promise(r => Office.context.mailbox.item.saveAsync(res => r(res.value)));
  
  const globalID = await new Promise(resolve => {
    Office.context.mailbox.item.getAllInternetHeadersAsync(result => {
      const headers = result.value || {};
      resolve(headers["UID"] || headers["vcal-uid"] || entryID); 
    });
  });

  const payload = {
    EntryID: entryID,
    globalID: globalID,
    Organizer: data.organizer,
    AuthorAlias: aliasStr,
    AuthorFirstname: firstNameStr,
    AuthorLastname: lastNameStr,
    OwnerEmail: ownerEmail,
    Subject: customerEvent,
    Start: formatDate(data.start),
    End: formatDate(data.end),
    Location: Location,
    CreationTime: CreationTime,
    ActivityType: data.actType,
    EngagementType: engagementType,
    OnSite: OnSite,
    CustInteraction: CustInteraction,
    Clevel: Clevel,
    Note: cleanBody
  };
  
  return JSON.stringify(payload);
}

function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = `status-message ${type}`;
  statusDiv.style.display = 'block';
}

function closePane() {
  Office.context.ui.closeContainer();

}
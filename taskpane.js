Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    // Initialize MSAL and get an access token
    getAccessToken().then(() => {
      // Get the email subject and body
      Office.context.mailbox.item.subject.getAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById('title').value = result.value;
        }
      });

      Office.context.mailbox.item.body.getAsync('text', result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById('notes').value = result.value;
        }
      });

      // Populate planners and assignees
      populatePlanners();
      populateAssignees();
    }).catch(error => {
      console.error('Error getting access token:', error);
    });
  }
});

async function populatePlanners() {
  try {
    const planners = await fetchPlanners(); // Implement this function to fetch planners from Graph API
    const plannerSelect = document.getElementById('planner');
    plannerSelect.innerHTML = planners.map(planner => `<option value="${planner.id}">${planner.name}</option>`).join('');
  } catch (error) {
    console.error('Error populating planners:', error);
  }
}

async function populateAssignees() {
  try {
    const assignees = await fetchAssignees(); // Implement this function to fetch users from Graph API
    const assigneeSelect = document.getElementById('assignee');
    assigneeSelect.innerHTML = assignees.map(user => `<option value="${user.id}">${user.displayName}</option>`).join('');
  } catch (error) {
    console.error('Error populating assignees:', error);
  }
}

document.getElementById('taskForm').addEventListener('submit', async event => {
  event.preventDefault();

  const title = document.getElementById('title').value;
  const planner = document.getElementById('planner').value;
  const assignee = document.getElementById('assignee').value;
  const dueDate = document.getElementById('dueDate').value;
  const priority = document.getElementById('priority').value;
  const notes = document.getElementById('notes').value;

  try {
    await createPlannerTask(title, planner, assignee, dueDate, priority, notes);
    alert('Task created successfully!');
  } catch (error) {
    console.error('Error creating task:', error);
    alert('Failed to create task.');
  }
});

async function createPlannerTask(title, planner, assignee, dueDate, priority, notes) {
  const accessToken = await getAccessToken();
  const task = {
    title,
    planId: planner,
    assignee: assignee,
    dueDateTime: dueDate,
    priority: priority,
    notes: notes
  };

  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/planner/tasks', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(task)
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    return response.json();
  } catch (error) {
    console.error('Error creating task:', error);
    throw error;
  }
}

// MSAL Configuration
const msalConfig = {
  auth: {
    clientId: 'ee62763c-3386-465c-8229-0e0b69447205', 
    authority: 'https://login.microsoftonline.com/c63f10b3-d204-4e0f-b951-7463aa432e51', 
    redirectUri: 'https://localhost:3000'
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: true
  }
};

// Create a new MSAL PublicClientApplication instance
const msalInstance = new Msal.PublicClientApplication(msalConfig);

async function getAccessToken() {
  const loginRequest = {
    scopes: ['https://graph.microsoft.com/.default']
  };

  try {
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    console.log('Access Token:', loginResponse.accessToken); // Log access token for debugging
    return loginResponse.accessToken;
  } catch (error) {
    console.error('Error during authentication:', error); // Log error
    throw error;
  }
}
  
  try {
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    return loginResponse.accessToken;
  } catch (error) {
    console.error('Error during authentication:', error);
    throw error;
  }
}

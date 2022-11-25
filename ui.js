async function displayUI() {    
    await signIn();

    // Display info from user profile
    const user = await getUser();
    var userName = document.getElementById('userName');
    userName.innerText = user.displayName;  

    // Hide login button and initial UI
    var signInButton = document.getElementById('signin');
    signInButton.style = "display: none";
    var content = document.getElementById('content');
    content.style = "display: block";

    var showPhotoButton= document.getElementById('showProfilePhoto');
    showPhotoButton.style = "display: block";

    var btnShowEvents = document.getElementById('btnShowEvents');
    btnShowEvents.style = "display: block";

      // beginning of function omitted for brevity
    content.style = "display: block";
    displayFiles();
}

async function displayProfilePhoto() {
    const userPhoto = await getUserPhoto();
    if (!userPhoto) {
        return;
    }
  
    //convert blob to a local URL
    const urlObject = URL.createObjectURL(userPhoto);
    // show user photo
    const userPhotoElement = document.getElementById('userPhoto');
    userPhotoElement.src = urlObject;
    var showPhotoButton= document.getElementById('showProfilePhoto');
    showPhotoButton.style = "display: none";
    var imgPhoto= document.getElementById('userPhoto');
    imgPhoto.style = "display: block";
}

//https://learn.microsoft.com/en-us/training/modules/msgraph-show-user-emails/6-exercise-show-users-email
var nextLink;
async function displayEmail() {
  var emails = await getEmails(nextLink);
  if (!emails || emails.value.length < 1) {
    return;
  }
  nextLink = emails['@odata.nextLink'];

  document.getElementById('displayEmail').style = 'display: none';

  var emailsUl = document.getElementById('emails');
  emails.value.forEach(email => {
    var emailLi = document.createElement('li');
    emailLi.innerText = `${email.subject} (${new Date(email.receivedDateTime).toLocaleString()})`;
    emailsUl.appendChild(emailLi);
  });
  window.scrollTo({ top: emailsUl.scrollHeight, behavior: 'smooth' });

  if (nextLink) {
    document.getElementById('loadMoreContainer').style = 'display: block';
  }
}

//https://learn.microsoft.com/en-us/training/modules/msgraph-access-user-events/6-exercise-configure-app-access-events
async function displayEvents() {
    var events = await getEvents();
    if (!events || events.value.length < 1) {
      var content = document.getElementById('content');
      var noItemsMessage = document.createElement('p');
      noItemsMessage.innerHTML = `No events for the coming week!`;
      content.appendChild(noItemsMessage)
  
    } else {
      var wrapperShowEvents = document.getElementById('eventWrapper');
      wrapperShowEvents.style = "display: block";
      const eventsElement = document.getElementById('events');
      eventsElement.innerHTML = '';
      events.value.forEach(event => {
        var eventList = document.createElement('li');
        eventList.innerText = `${event.subject} - From  ${new Date(event.start.dateTime).toLocaleString()} to ${new Date(event.end.dateTime).toLocaleString()} `;
        eventsElement.appendChild(eventList);
      });
    }
    var btnShowEvents = document.getElementById('btnShowEvents');
    btnShowEvents.style = "display: none";
  }

//https://learn.microsoft.com/en-us/training/modules/msgraph-manage-files/6-exercise-display-user-files
async function displayFiles() {
    const files = await getFiles();
    const ul = document.getElementById('downloadLinks');
    while (ul.firstChild) {
      ul.removeChild(ul.firstChild);
    }
    for (let file of files) {
      if (!file.folder && !file.package) {
        let a = document.createElement('a');
        a.href = '#';
        a.onclick = () => { downloadFile(file); };
        a.appendChild(document.createTextNode(file.name));
        let li = document.createElement('li');
        li.appendChild(a);
        ul.appendChild(li);
      }
    }
  }

//https://learn.microsoft.com/en-us/training/modules/msgraph-manage-files/10-exercise-upload-user-files
  function fileSelected(e) {
    displayUploadMessage(`Uploading ${e.files[0].name}...`);
    uploadFile(e.files[0])
    .then((response) => {
      displayUploadMessage(`File ${response.name} of ${response.size} bytes uploaded`);
      displayFiles();
    });
  }
  
  function displayUploadMessage(message) {
      const messageElement = document.getElementById('uploadMessage');
      messageElement.innerText = message;
  }  
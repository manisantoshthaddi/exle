
// var queryString = window.location.search;
// var urlParams = new URLSearchParams(queryString);
// var dataReceived = urlParams.get('data');
// console.log(dataReceived)
const schemaList = document.getElementById('schemaList');
let spinnerEl = document.getElementById("spinner");
spinnerEl.classList.remove("d-none");
// window.addEventListener('scroll', () => {
//   const newHeight = window.innerHeight - window.scrollY;
//   console.log(newHeight)
//   schemaList.style.height = `${newHeight}px`;
// });

var schemanames;
var myModal = new bootstrap.Modal(document.getElementById("myModal"));
const queryString = window.location.search;
const urlParams = new URLSearchParams(queryString);
const dataParameter = urlParams.get('data');
//const dataforviewfava = JSON.parse(decodeURIComponent(dataParameter));
console.log(dataParameter)
const jsonStr = atob(dataParameter);
const dataforviewfava = JSON.parse(jsonStr)

const ti = dataforviewfava.ti;
const iu = dataforviewfava.iu;
const token = dataforviewfava.token
//console.log(`ti: ${ti}, iu: ${iu}, token: ${token}`);
//const schemaList = document.getElementById("schemaList");
const searchBox = document.getElementById("searchInput");
const QuerryText = document.getElementById("UserQuery")

// searchBox.addEventListener("input", function () {
//   const query = searchBox.value;
//   updateResults(query);
// });

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    //var data = Office.context.getSettings('data');
    
    // OfficeExtension.ExtensionHelpers.getCustomDataValue('data', function(data) {
    //   if (data) {
    //       console.log('Received data: ' + data);
    //   } else {
    //       console.log('Data not found');
    //   }
    // });
    
    document.getElementById("Add").blur();
    // Add Favorite Button
    document.getElementById("Add").onclick = AddFav;
    //console.log(window);
    document.getElementById('searchInput').addEventListener('input', function () {
      const query = searchBox.value;
      updateResults(query);
    });
    document.getElementById('refresh').addEventListener('click',function(){
      //console.log('mani');
      //document.getElementById('schemaList').style.display = 'none';
      spinnerEl.classList.remove("d-none");
      schemaList.textContent = "";
      dataCatalogue();
    })
    document.getElementById('Editbutton').addEventListener('click',function(){
      //console.log('mani'+document.getElementById('UserQuery').value);
      let querytext= document.getElementById('UserQuery').value
      //console.log(querytext)
      let messageObject = { messageType: "queryname", query:querytext};
      let jsonMessage = JSON.stringify(messageObject);
      Office.context.ui.messageParent(jsonMessage);
      
    })
  }
});

const fetchData = async (url) => {
  const res = await fetch(url, {
    method: 'GET',
    headers: { 'Authorization': 'Bearer ' + token },
  });
  return await res.json();
};
const dataCatalogue = () => {
  //console.log("manisantosh")
  const url = `${iu}/${ti}/IONSERVICES/datacatalog/v1/object/list?type=JSON,DSV`;
  fetchData(url)
      .then((json) => {
        try{
          const obj = json.objects.filter((obj) => obj.subType === 'JSONStream' ||  obj.type === 'DSV');
          schemanames = obj.map((obj)=>(obj.name))
          updateSchemaList(schemanames);
        }
        catch{
          document.getElementById("modalHeading").innerHTML = "Sign Out",
          document.getElementById("modalText").innerHTML = "Signed out Successfully",
          myModal.show()
        }
      })
      // .catch(error =>
      //   document.getElementById("modalHeading").innerHTML = "Sign Out",
      //   document.getElementById("modalText").innerHTML = "Signed out Successfully",
      //   myModal.show()
      // )
};
if(ti!=null && iu!=null && token!=null){
  dataCatalogue()
}
else{
  console.log("somthing went worng")
}

function updateSchemaList(schemanames) {
  schemaList.textContent= "";
  spinnerEl.classList.add("d-none");
  schemanames.forEach(schemaName => {
    const innerDiv = document.createElement("div");
    innerDiv.className = "schema-item";
    innerDiv.textContent = schemaName;
    innerDiv.addEventListener("click", () => compassQuery(schemaName));
    schemaList.appendChild(innerDiv);
  });
}

function updateResults(query) {
  const filteredSchemanames = schemanames.filter(schemaName => schemaName.toLowerCase().includes(query.toLowerCase()));
  updateSchemaList(filteredSchemanames);
}

const compassQuery = (name) => {
      const url = `${iu}/${ti}/IONSERVICES/datacatalog/v1/object/${name}`;
      fetchData(url)
          .then((json) => {
              try{
              const keys = Object.keys(json.schema.properties);
              const query = keys.map((key) => `${name}.${key}`).join(',\n ');
              QuerryText.value = `Select ${query} From "${name}"`
              var messageObject = { messageType: "queryname", query: `Select ${query} From "${name}"`};
              var jsonMessage = JSON.stringify(messageObject);
              Office.context.ui.messageParent(jsonMessage);
              }
              catch{
                document.getElementById("modalHeading").innerHTML = "Sign Out",
                document.getElementById("modalText").innerHTML = "Signed out Successfully",
                myModal.show()
              }
          })
          .catch((err) => console.log(err))
          .finally(() => {
          });
};

// Function to display table 
function displayTable(tbody, serialNo, queryTag, timestamp, acceptSpan, deleteSpan) {
  let row = tbody.insertRow();
  let cell1 = row.insertCell(0);
  let cell2 = row.insertCell(1);
  let cell3 = row.insertCell(2);
  let cell4 = row.insertCell(3);
  let cell5 = row.insertCell(4);

  // Set First Cell Value
  cell1.innerText = serialNo;

  // Set Second Cell
  cell2.appendChild(queryTag);

  // Set Third Cell Value
  cell3.innerText = timestamp;

  // Set fourth cell value
  cell4.appendChild(acceptSpan);

  // Set fifth cell value
  cell5.appendChild(deleteSpan);
}

//Function to display UI
function generateTable() {
  // Display table
  document.querySelector('.table').style.display = 'table';

  // Select the existing table element
  let tbody = document.querySelector('.table_body');

  // Set empty table body
  let newTbody = document.createElement('tbody');
  newTbody.classList.add('table_body');

  // Initialize localStorage if empty
  if (localStorage.getItem('queryList') === null) {
    localStorage.setItem('queryList', JSON.stringify([]));
    console.log(localStorage.queryList);
  }

  // If there are no queries saved make the table display as none
  if (JSON.parse(localStorage.getItem('queryList')).length === 0)
    // Hide table columns
    document.querySelector('.table').style.display = 'none';

  // Set rows
  JSON.parse(localStorage.getItem('queryList')).forEach(([query, timestamp], seno) => {
    //console.log(query, timestamp);

    // Set Query Field
    let preTag = document.createElement("pre");
    preTag.style.marginBottom = '0px';
    //console.log(sqlFormatter.format(query, { language: 'mysql' }));

    // Format query tag
    let code = document.createElement('code');
    code.setAttribute('query-content', `${query}`);
    code.textContent = `${sqlFormatter.format(query, { language: 'mysql', "tabWidth": 3, "keywordCase": "upper" })}`;
    preTag.appendChild(code);
    //preTag.innerHTML = `<code query-content="${query}">${format(query, { language: 'mysql' })}</code>`;

    // Set Accept Button
    let acceptSpan = document.createElement('span');
    acceptSpan.setAttribute('class', 'open');
    acceptSpan.setAttribute('title', 'Select Query');
    acceptSpan.innerHTML = `<img width="16" height="16" src="https://localhost:3000/assets/accept.png" alt="Infor" />`;

    // Add Event Listener for select span button
    acceptSpan.onclick = function () {
      setButtonEventListener(sqlFormatter.format(query, { language: 'mysql', "keywordCase": "upper" }).replace(/\n/g, " "));;
    };

    // Set Delete Button
    let deleteSpan = document.createElement('span');
    deleteSpan.setAttribute('class', 'close');
    deleteSpan.setAttribute('title', 'Delete Query');
    deleteSpan.innerHTML = `<img width="16" height="16" src="https://localhost:3000/assets/remove.png" alt="Infor" />`;
    // Add Event Listener for delete span button
    deleteSpan.onclick = function () {
      setCloseEventListener(this);
    };

    // Update Table UI
    displayTable(newTbody, seno + 1, preTag, timestamp, acceptSpan, deleteSpan);

  });

  // Replace old tBody with new one
  //console.log(newTbody);
  tbody.replaceWith(newTbody);

}

window.addEventListener('load', () => {
  generateTable();

});

function setButtonEventListener(opt) {
  //console.log(opt);
  var messageObject = { messageType: "UserQuery", query: opt };
  var jsonMessage = JSON.stringify(messageObject);
  Office.context.ui.messageParent(jsonMessage);

}

//Set Event Listener to Remove ION API File from dropdown
function setCloseEventListener(row) {
  //console.log(row);

  // select element with queryText
  let i = row.parentNode.parentNode.rowIndex;
  //console.log(i);
  let tr = document.getElementsByTagName("tr")[i];

  let name = tr.cells[1].firstChild.firstChild.getAttribute('query-content');
  //console.log(name);
  let queryList = JSON.parse(localStorage.getItem('queryList'));

  //console.log(queryList.filter(([query]) => query !== name));
  localStorage.setItem('queryList', JSON.stringify(queryList.filter(([query]) => query !== name)));

  generateTable();
}

// Function to add favorite element
function AddFav() {
  // Blur the button
  document.getElementById("Add").blur();

  let tbody = document.querySelector('.table_body');

  // Read Query
  let query = document.getElementById('UserQuery').value.trim();

  // Original unFormatted query for message
  //let originalQuery = query;

  // format query using library to avoid duplication
  query = sqlFormatter.format(query, { language: 'mysql', "tabWidth": 3, "keywordCase": "upper" });
  //console.log(query);

  // Check if empty query is entered
  if (query === "") {
    let myModal = new bootstrap.Modal(document.getElementById("myModal"));
    document.getElementById("modalHeading").innerHTML = "Empty Query";
    document.getElementById("modalText").innerHTML = "Please enter a Query";
    myModal.show();
  }

  // Check if Query already exists;

  else if (localStorage.getItem('queryList') && JSON.parse(localStorage.queryList).map(([query]) => query).includes(query)) {
    //console.log("Same query");
    let myModal = new bootstrap.Modal(document.getElementById("myModal"));
    document.getElementById("modalHeading").innerHTML = "Duplicate Query";
    document.getElementById("modalText").innerHTML = "The entered query matches an existing query, please verify.";
    myModal.show();
  }

  else {
    // Display table columns
    document.querySelector('.table').style.display = 'table';

    // Check if queryList property exists in localstorage
    if (localStorage.getItem('queryList') === null) {
      localStorage.setItem('queryList', JSON.stringify([]));
      //console.log(localStorage.queryList);
    }

    // Check if 20 queries are already saved
    else if (JSON.parse(localStorage.getItem('queryList')).length === 20) {
      //console.log("Limit reached");
      let myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Query Limit";
      document.getElementById("modalText").innerHTML = "Saved Query Limit is 20, please verify.";
      myModal.show();
    }

    // Valid query. Add to localstorage
    else {
      // Add query and timestamp to queryList
     // let timestamp = new Date().toLocaleString();
      
      // Set queryList to localStorage
      let queryList = JSON.parse(localStorage.getItem('queryList'));
      let name = document.getElementById("queryname").value || "Querry_"+(parseInt(queryList.length)+1)
      //console.log(queryList);
      queryList.push([query, name]);
      localStorage.setItem('queryList', JSON.stringify(queryList));

      // Make query textbox as null
      document.getElementById('UserQuery').value = "";

      // Set Query Field
      let preTag = document.createElement("pre");
      preTag.style.marginBottom = '0px';
      preTag.classList.add('queryOption');
      //console.log(sqlFormatter.format(query, { language: 'mysql' }));

      // Format query tag
      let code = document.createElement('code');
      code.setAttribute('query-content', `${query}`);
      code.textContent = `${sqlFormatter.format(query, { language: 'mysql', "tabWidth": 3, "keywordCase": "upper" })}`;
      preTag.appendChild(code);
      //preTag.innerHTML = `<code query-content="${query}">${format(query, { language: 'mysql' })}</code>`;

      // Set Accept Button
      let acceptSpan = document.createElement('span');
      acceptSpan.setAttribute('class', 'open');
      acceptSpan.setAttribute('title', 'Select Query');
      acceptSpan.innerHTML = `<img width="16" height="16" src="https://localhost:3000/assets/accept.png" alt="Infor" />`;

      // Add Event Listener for select span button
      acceptSpan.onclick = function () {
        setButtonEventListener(sqlFormatter.format(query, { language: 'mysql', "keywordCase": "upper" }).replace(/\n/g, " "));
      };

      // Set Delete Button
      let deleteSpan = document.createElement('span');
      deleteSpan.setAttribute('class', 'close');
      deleteSpan.setAttribute('title', 'Delete Query');
      deleteSpan.innerHTML = `<img width="16" height="16" src="https://localhost:3000/assets/remove.png" alt="Infor" />`;
      // Add Event Listener for delete span button
      deleteSpan.onclick = function () {
        setCloseEventListener(this);
      };

      // Update Table UI
      displayTable(tbody, queryList.length, preTag, name, acceptSpan, deleteSpan);
    }

  }
}






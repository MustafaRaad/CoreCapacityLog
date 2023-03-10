import React, { Component } from 'react';
import hello from 'hellojs/dist/hello.all.js';
import axios from 'axios';
import { JsonToExcel } from "react-json-excel";
import * as htmlToImage from 'html-to-image';

class Home extends Component {
  constructor(props) {
    super(props);

    const msft = hello('msft').getAuthResponse();

    this.state = {
      userData: "",
      apiData: {},
      toExportData: [
      ],
      date: '',
      time: '',
      token: msft.access_token
    };

    // this.onFetchExcelData = this.onFetchExcelData.bind(this);
    this.onLogout = this.onLogout.bind(this);
    this.createImg = this.createImg.bind(this);

  }
  componentDidMount() {
    // Set the execution time to 10:00 AM every day
    const executionTime = new Date();
    executionTime.setHours(12);
    executionTime.setMinutes(0);
    executionTime.setSeconds(0);

    // Calculate the interval between now and the execution time
    const interval = executionTime.getTime() - Date.now();
    if (interval > 0) {
      // Set an interval to execute the task every day at the specified time
      this.interval = setInterval(() => {
        console.log('Task executed! foog', interval);
      }, interval);
    }
  }
  createImg() {
    const element = document.querySelector('#dataTable');
    htmlToImage.toPng(element)
      .then((dataUrl) => {
        // console.log('excuted img');
        // download(dataUrl, 'my-node.png');
        var img = new Image();
        img.src = dataUrl;
        document.body.appendChild(img);
      }).catch(function (error) {
        console.error('oops, something went wrong!', error);
      });

    // axios.post('https://hooks.slack.com/services/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', { image: this.dataUrl  }, {headers: {"content-type": "application/json"}}).then((response) => {
    //   console.log(response);
    //   console.log(response.data);
    // });
    // const apiPostUrl = 'https://hooks.slack.com/services/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'
    const apiPostUrl = 'https://slack.com/api/auth.test'
    const apiPostToken = 'Bearer xapp-1-A04BRV40YNT-4583645743618-dd429ea159a03517209d41285427a77140ad131bc5fb4afaa9d353127ea3ef44'

    // const apiPostBody = {
    //   title: "My first Slack Message",
    //   text: "Random example message text",
    //   color: "#00FF00",
    // };

    const body = `Token=${apiPostToken}`
    axios.post(apiPostUrl, {
      // channel: '#general',
      text: 'Hello, world!'
    }, {
      headers: {
        'Authorization': `${apiPostToken}`
      }
    }).then(res => {
      console.log(res)
    }).catch(err => {
      console.log(err)
    })

    
  }

  async componentWillMount() {
    clearInterval(this.interval);
    const { token } = this.state;
    axios.get(
      'https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName',
      { headers: { Authorization: `Bearer ${token}` } }
    ).then(res => {
      const userData = res.data;
      this.setState({ userData });
    });
    axios
      .get("https://graph.microsoft.com/v1.0/drives/b!rv9qh9qe60ay10NPvtjGbxGy1-IwSUdCgPhiZyJfMXjt8D0-O3QQR5DYvhLdsgis/items/01A4OPHAHO5KQRHTRJPRA3MKJP4SXW2TEY/workbook/worksheets('Capacity tracking ')/usedRange",
        { headers: { Authorization: `Bearer ${token}` } }
      )
      .then(res => {
        const apiData = res.data;
        this.setState({ apiData });
      }).then(() => {
        this.createImg()
      }
      )
      .catch(err => console.error(err));
  }
  onLogout() {
    hello('msft').logout().then(
      () => this.props.history.push('/'),
      e => console.error(e.error.message)
    );
  }

  renderUserData() {
    const { userData } = this.state;
    const myEmailAddress = userData.mail || userData.userPrincipalName;
    const userViewData = <tr>
      <td>{userData.displayName}</td>
      <td>{myEmailAddress}</td>
    </tr>
    return (

      <table className="table stripped">
        <thead>
          <tr>
            <th>Name</th>
            <th>Email</th>
          </tr>
        </thead>
        <tbody>
          {userViewData}
        </tbody>
      </table>
    );
  }
  renderApiData() {
    const { apiData, toExportData } = this.state;
    if (Object.keys(apiData).length > 0) {
      const apiTextValuesData = apiData.text;
      const firstRowData = apiTextValuesData[1];
      var temp = '';
      // getting last index to render last date
      for (let index = firstRowData.length; index > 0; index--) {
        const element = firstRowData[index];
        if (element !== '' && element !== undefined) {
          // console.table(element, index);
          this.lastColIndex = index;
          break
        }
      }
      const dataHeader = <thead className='table-primary'>
        <tr>
          <th>{firstRowData[0]}</th>
          <th>{firstRowData[1]}</th>
          <th>{firstRowData[this.lastColIndex]}</th>
        </tr>
      </thead>
      const dataTable = <tbody>
        {
          apiTextValuesData.map((element, index) => {
            if (!(element[0] === '' && element[1] === '' && element[this.lastColIndex] === '')) {
              toExportData.push({ 1: element[0], 2: element[1], 3: element[this.lastColIndex] })
              element[0] === '' ? element[0] = temp : temp = element[0];
              return (
                <tr key={index}>
                  <td>{element[0]}</td>
                  <td>{element[1]}</td>
                  <td>{element[this.lastColIndex]}</td>
                </tr>
              );
            } else {
              return false
            }
          })}
      </tbody>

      return (
        <table className="table table-striped table-borderless border" id='dataTable'>
          {dataHeader}
          {dataTable}
        </table>
      )
    } else {
      return "Loading"
    }

  }

  render() {
    let today = new Date();
    const date = today.getDate() + '-' + (today.getMonth() + 1) + '-' + today.getFullYear();
    const time = today.getHours() + ":" + today.getMinutes()

    const refresh = () => { window.location.reload() }
    const filename = "Core Capacity Log",
      fields = {
        1: "1",
        2: "2",
        3: "3"
      }
    console.log(this.state.toExportData);
    return (
      <div className='container'>

        {this.renderUserData()}
        <div className='row'>
          <div className='col-2'>
            {/* <button onClick={this.componentWillMount.createImg()} className="btn btn-light">Refresh Pageaa</button> */}
          </div>
          <div className='col-1'>
            <button onClick={this.onLogout} className="btn btn-light">Logout</button>
          </div>
          <div className='col-2'>
            <button onClick={refresh} className="btn btn-light">Refresh Page</button>
          </div>
          <div className='col-2'>
            <JsonToExcel
              data={this.state.toExportData}
              className="btn btn-success btn-large"
              filename={filename}
              fields={fields}
              text="Download Excel"
            />
          </div>
          <div className='col-3'>
            <p>last Update in: {date} / {time} </p>
          </div>
        </div>
        <div id='dataTable' style={{ backgroundColor: "#fff" }}>
          {this.renderApiData()}
        </div>
      </div>
    );
  }
}

export default Home;

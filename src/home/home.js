import React, { Component } from 'react';
import hello from 'hellojs/dist/hello.all.js';
import axios from 'axios';

class Home extends Component {
  constructor(props) {
    super(props);

    const msft = hello('msft').getAuthResponse();

    this.state = {
      userData: "",
      successMessage: "",
      apiData: {},
      date: '',
      time: '',
      token: msft.access_token
    };

    // this.onFetchExcelData = this.onFetchExcelData.bind(this);
    this.onLogout = this.onLogout.bind(this);


  }
  async componentDidMount() {
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
    const { apiData } = this.state;
    if (Object.keys(apiData).length > 0) {
      const apiTextValuesData = apiData.text;
      const firstRowData = apiTextValuesData[1];
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
            return (
              <tr key={index}>
                <td>{element[0]}</td>
                <td>{element[1]}</td>
                <td>{element[this.lastColIndex]}</td>
              </tr>
            );
          })}
      </tbody>

      return (
        <table className="table table-striped table-borderless border">
          {dataHeader}
          {dataTable}
        </table>
      )
    } else {
      return "Loading"
    }

  }

  render() {
    var today = new Date();
    const date = (today.getMonth() + 1) + '-' + today.getDate();
    const time = today.getHours() + ":" + today.getMinutes()
    
    const refresh = () => { window.location.reload() }
    
    return (
      <div className='container'>
        {this.renderUserData()}
        <div className='row'>
          <div className='col-1'>
            <button onClick={this.onLogout} className="btn btn-light">Logout</button>
          </div>
          <div className='col-2'>
            <button onClick={refresh} className="btn btn-light">Refresh Page</button>
          </div>
          <div className='col-3'>
            <p>last Update in: {date} / {time} </p>
          </div>
        </div>
        {this.renderApiData()}
      </div>
    );
  }
}

export default Home;

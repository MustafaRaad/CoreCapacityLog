import React, { Component } from 'react';
import hello from 'hellojs/dist/hello.all.js';
import axios from 'axios';

class Home extends Component {
  constructor(props) {
    super(props);

    const msft = hello('msft').getAuthResponse();

    this.state = {
      me: "",
      successMessage: "",
      xlsxData: {},
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
      const me = res.data;
      this.setState({ me });
    });
    axios
      .get("https://graph.microsoft.com/v1.0/drives/b!rv9qh9qe60ay10NPvtjGbxGy1-IwSUdCgPhiZyJfMXjt8D0-O3QQR5DYvhLdsgis/items/01A4OPHAHO5KQRHTRJPRA3MKJP4SXW2TEY/workbook/worksheets('Capacity tracking ')/usedRange",
        // { index: null, values },
        { headers: { Authorization: `Bearer ${token}` } }
      )
      .then(res => {
        // console.log(res, "API Call");
        const xlsxData = res.data;

        // const successMessage = "Successfully wrote your data to demo.xlsx!";
        this.setState({ xlsxData });
        // const customHeadings = res.data.map(item => ({
        //   "Article Id": item.id,
        //   "Article Title": item.title
        // }))

        // console.log(customHeadings);
        // setData(customHeadings)
      })
      .catch(err => console.error(err));
  }

  onLogout() {
    hello('msft').logout().then(
      () => this.props.history.push('/'),
      e => console.error(e.error.message)
    );
  }

  renderMe() {
    const { me } = this.state;
    const myEmailAddress = me.mail || me.userPrincipalName;

    return (
      <tr>
        <td>{me.displayName}</td>
        <td>{myEmailAddress}</td>
      </tr>
    );
  }
  renderMe2() {
    const { xlsxData } = this.state;
    if (Object.keys(xlsxData).length > 0) {
      const valuesData = xlsxData.text;
      const valueTypes = xlsxData.valueTypes;
      var lastColIndex = 0
      // console.table(valuesData[1], "xlsxData renderMe2");
      for (let index = valuesData[1].length; index > 0; index--) {
        const element = valuesData[1][index];
        if (element !== '' && element !== undefined) {
          console.table(element, index);
          this.lastColIndex = index;
          break
        }
      }

      // valuesData[1].forEach((element, index) => {
      //   if (valueTypes[1][index] !== 'Empty') {
      //     // console.table(element);
      //     // this.lastCol.push(element);
      //     console.table(element);
      //   }
      //   // console.log(element[element.length-10]);
      // });
      // console.log(valuesData, "renderMe2");
      // console.log(xlsxData.valueTypes[0][0] == 'Empty' && xlsxData.valueTypes[0][1] == 'Empty' && xlsxData.valueTypes[0][2] == 'Empty');
      return (
        // {/* <td>{valuesData}</td> */ }
        //   {
        //   valuesData.forEach(element => {
        //     console.log(element);
        //   })
        // }
        // {if (xlsxData.valueTypes.element)}
        <tbody>
          {
            valuesData.map((element, index) => {
              return (
                <tr key={index}>
                  <td>{element[0]}</td>
                  <td>{element[1]}</td>
                  <td>{element[this.lastColIndex]}</td>
                </tr>
              );
            })}
        </tbody>
      )
    } else {
      return "Loading"
    }

  }

  render() {
    return (
      <div className='container'>
        <table className="table stripped">
          <thead>
            <tr>
              <th>Name</th>
              <th>Email</th>
            </tr>
          </thead>
          <tbody>
            {this.renderMe()}
          </tbody>
        </table>
        <table className="table table-striped-columns ">
          <thead>
            <tr>
              <th>Site</th>
              <th>Supplier</th>
              <th>Supplier</th>
            </tr>
          </thead>
          {this.renderMe2()}
        </table>

        {/* <button onClick={this.onFetchExcelData}>Write to Excel</button> */}
        <button onClick={this.onLogout}>Logout</button>
        {/* <p>{this.state.successMessage}</p> */}
      </div>
    );
  }
}

export default Home;

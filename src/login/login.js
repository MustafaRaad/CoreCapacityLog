import React, { Component } from 'react';
import hello from 'hellojs/dist/hello.all.js';

import { Configs } from '../configs';

class Login extends Component {
  constructor(props) {
    super(props);

    this.onLogin = this.onLogin.bind(this);
  }

  onLogin() {
    hello('msft').login({ scope: Configs.scope }).then(
      () => this.props.history.push('/home'),
      e => console.error(e.error.message)
    );
  }

  render() {
    return (
      <div className='container vh-100 d-flex justify-content-center align-items-center'>
        <button onClick={this.onLogin} className="btn btn-primary btn-lg">Sign in with your Microsoft account</button>
      </div>
    );
  }
}

export default Login;

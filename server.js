const express = require('express');
const cors = require('cors');
const app = express();

app.use(cors());

app.get('/api/:method', (req, res) => {
  const method = req.params.method;
  const url = `https://slack.com/api/${method}`;

  fetch(url, {
    headers: {
      'Authorization': `Bearer ${process.env.SLACK_TOKEN}`,
      'Content-Type': 'application/x-www-form-urlencoded'
    }
  })
    .then(response => response.json())
    .then(data => res.send(data))
    .catch(error => console.error(error));
});

app.listen(3000, () => console.log('CORS proxy listening on port 3000'));

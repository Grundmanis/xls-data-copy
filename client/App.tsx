
import React from "react";

const App: React.FC = () => {
  return <>
  <form action="/4fG7hJkLmN8pQrStUvWx/upload" method="POST" encType="multipart/form-data">
    <label>From File:</label>
    <input type="file" name="from" id="from" required />
    
    <label>To File:</label>
    <input type="file" name="to" id="to" required />
    
    <label>Partners file:</label>
    <input type="file" name="partners" id="partners" required />

    <button type="submit">Upload Files</button>
  </form>
  </>;
};

export default App;
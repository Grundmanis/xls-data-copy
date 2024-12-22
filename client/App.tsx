
import React from "react";

const App: React.FC = () => {
  return <>
  <form action="/upload" method="POST" enctype="multipart/form-data">
    <label for="from">From File:</label>
    <input type="file" name="from" id="from" required />
    
    <label for="to">To File:</label>
    <input type="file" name="to" id="to" required />

    <label for="targetTab">Target tab:</label>
    <select name="targetTab" id="targetTab">
      <option selected value="Dangerous_Cargo">Dangerous_Cargo</option>
      <option value="Cargo">Cargo</option>
    </select>
    
    <button type="submit">Upload Files</button>
  </form>
  </>;
};

export default App;
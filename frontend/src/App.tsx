import React from 'react';
import './App.css';
import UploadPage from './pages/UploadPage';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>Word文档格式优化器</h1>
      </header>
      <main>
        <UploadPage />
      </main>
    </div>
  );
}

export default App;
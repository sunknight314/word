import React from 'react';
import './JsonDisplay.css';

interface JsonDisplayProps {
  data: any;
  title: string;
}

const JsonDisplay: React.FC<JsonDisplayProps> = ({ data, title }) => {
  return (
    <div className="json-display">
      <h3>{title}</h3>
      <pre className="json-content">
        {JSON.stringify(data, null, 2)}
      </pre>
    </div>
  );
};

export default JsonDisplay;
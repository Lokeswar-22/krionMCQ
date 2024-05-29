import React, { useState, useEffect } from 'react';
import axios from 'axios';
import './App.css';

function App() {
  const [candidateName, setCandidateName] = useState('');
  const [candidateRegisterNumber, setCandidateRegisterNumber] = useState('');
  const [isValidCandidate, setIsValidCandidate] = useState(false);
  const [currentQuestionId, setCurrentQuestionId] = useState(1);
  const [question, setQuestion] = useState('');
  const [options, setOptions] = useState([]);
  const [selectedOption, setSelectedOption] = useState('');
  const [resultMessage, setResultMessage] = useState('');
  const [candidateInfo, setCandidateInfo] = useState([]);

  useEffect(() => {
    fetchCandidateInfo();
  }, []);

  const fetchCandidateInfo = async () => {
    try {
      if (candidateName === 'admin' && candidateRegisterNumber === 'adminhr123') {
        setIsValidCandidate(true);
        const response = await axios.get('http://localhost:3000/candidateinfo');
        setCandidateInfo(response.data);
      } else {
        const response = await axios.get('http://localhost:3000/candidateinfo');
        const candidate = response.data.find(candidate => candidate.candidateName === candidateName && candidate.candidateRegisterNumber === candidateRegisterNumber);
        setIsValidCandidate(!!candidate);
      }
    } catch (error) {
      console.error('Error fetching candidate info:', error);
    }
  };

  const fetchQuestion = async () => {
    try {
      const response = await axios.get(`http://localhost:3000/questions/${currentQuestionId}`);
      setQuestion(response.data.question);
      setOptions(response.data.options);
    } catch (error) {
      console.error('Error fetching question:', error);
    }
  };

  const submitAnswer = async () => {
    try {
      const response = await axios.post(`http://localhost:3000/answer/${currentQuestionId}`, {
        answer: selectedOption,
        candidateName: candidateName
      });
      setResultMessage(response.data);
      setSelectedOption('');

      if (currentQuestionId < 20) {
        setCurrentQuestionId(prevId => prevId + 1);
      } else {
        setResultMessage(`Successfully completed the test ${candidateName}`);
      }
    } catch (error) {
      console.error('Error submitting answer:', error);
    }
  };

  const handleLogin = () => {
    fetchCandidateInfo();
  };

  const handleOptionChange = event => {
    setSelectedOption(event.target.value);
  };

  useEffect(() => {
    if (isValidCandidate) {
      fetchQuestion();
    }
  }, [isValidCandidate, currentQuestionId]);

  return (
    <div className="App">
      <header className="App-header">
        <h1>Krion Consulting</h1>
      </header>
      {!isValidCandidate ? (
        <div className="login-container">
          <h2>Login</h2>
          <input type="text" placeholder="Candidate Name" value={candidateName} onChange={e => setCandidateName(e.target.value)} />
          <input type="text" placeholder="Register Number" value={candidateRegisterNumber} onChange={e => setCandidateRegisterNumber(e.target.value)} />
          <button onClick={handleLogin}>Login</button>
        </div>
      ) : (
        <div className="question-container">
          {candidateName === 'admin' && candidateRegisterNumber === 'adminhr123' ? (
            <div className="candidate-info">
              <h2>Candidate Info</h2>
              <table>
                <thead>
                  <tr>
                    <th>ID</th>
                    <th>Candidate Name</th>
                    <th>Register Number</th>
                    <th>Scored mark</th>
                  </tr>
                </thead>
                <tbody>
                  {candidateInfo.map(candidate => (
                    <tr key={candidate.id}>
                      <td>{candidate.id}</td>
                      <td>{candidate.candidateName}</td>
                      <td>{candidate.candidateRegisterNumber}</td>
                      <td>{candidate.totalMarks}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="question-card">
              <h2>Question {currentQuestionId}</h2>
              <h3>{question}</h3>
              <ul className="options-list">
                {options.map(option => (
                  <li key={option}>
                    <input
                      type="radio"
                      id={option}
                      name="option"
                      value={option}
                      checked={selectedOption === option}
                      onChange={handleOptionChange}
                    />
                    <label htmlFor={option}>{option}</label>
                  </li>
                ))}
              </ul>
              <button onClick={submitAnswer}>Submit Answer</button>
              {resultMessage && <p>{resultMessage}</p>}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export default App;

import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import Home from './components/Home';
import Navbar from './components/Navbar';
import HackathonRegistration from './components/HackathonRegistration';
import MasterclassRegistration from './components/MasterclassRegistration';
import GFGRegistration from './components/GFGRegistration';
import Background from './components/Background';
import { AnimationProvider } from './context/AnimationContext';
import './App.css';

function App() {
  return (
    <Router>
      <AnimationProvider>
        <div className="app-container">
          <Navbar />
          <Background />
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/register/hackathon" element={<HackathonRegistration />} />
            <Route path="/register/masterclass" element={<MasterclassRegistration />} />
            <Route path="/register/gfg" element={<GFGRegistration />} />
          </Routes>
        </div>
      </AnimationProvider>
    </Router>
  );
}

export default App;

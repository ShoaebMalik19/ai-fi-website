import React from 'react';
import { motion } from 'framer-motion';

const About = () => {
    return (
        <section className="section-container">
            <h2 className="text-gradient heading-md">
                About AIFi
            </h2>
            <div style={{
                display: 'flex',
                flexWrap: 'wrap',
                justifyContent: 'center',
                gap: '2rem'
            }}>
                <motion.div
                    className="glass-card neon-border"
                    initial={{ opacity: 0, x: -50 }}
                    whileInView={{ opacity: 1, x: 0 }}
                    transition={{ duration: 0.5 }}
                    style={{ flex: '1 1 300px', maxWidth: '600px', textAlign: 'left' }}
                >
                    <h3 style={{ color: 'var(--primary-cyan)', marginBottom: '1rem', fontSize: '1.5rem' }}>Vision</h3>
                    <p style={{ lineHeight: '1.6', fontSize: '1rem', color: '#e0e0e0' }}>
                        To empower a community of learners at REVA University with cutting-edge knowledge and practical skills in Artificial Intelligence and Machine Learning, cultivating innovation, collaboration, and leadership that bridges academic excellence with real-world technological impact.
                    </p>
                </motion.div>

                <motion.div
                    className="glass-card neon-border"
                    initial={{ opacity: 0, x: 50 }}
                    whileInView={{ opacity: 1, x: 0 }}
                    transition={{ duration: 0.5 }}
                    style={{ flex: '1 1 300px', maxWidth: '600px', textAlign: 'left' }}
                >
                    <h3 style={{ color: 'var(--secondary-purple)', marginBottom: '1rem', fontSize: '1.5rem' }}>Mission</h3>
                    <p style={{ lineHeight: '1.5', fontSize: '1rem', color: '#e0e0e0', marginBottom: '1rem' }}>
                        To foster a hands-on learning ecosystem in Artificial Intelligence and Machine Learning by organising workshops, projects, hackathons, research initiatives, and industry interactions—enabling students of REVA University to develop problem-solving skills, innovate responsibly, and become future-ready AI professionals.
                    </p>
                    <ul style={{ paddingLeft: '20px', lineHeight: '1.8', fontSize: '0.95rem', color: '#e0e0e0' }}>
                        <li><strong>Cultivate Knowledge:</strong> Provide a platform for students to explore and master AI & ML concepts through workshops, masterclasses, and hands-on projects.</li>
                        <li><strong>Foster Innovation:</strong> Encourage creative problem-solving and ideation to address societal and industrial challenges.</li>
                        <li><strong>Bridge the Gap:</strong> Connect students with industry experts, alumni, and opportunities to ensure industry readiness.</li>
                        <li><strong>Promote Collaboration:</strong> Build a strong community of learners and developers working together to push the boundaries of technology.</li>
                    </ul>
                </motion.div>
            </div>
        </section>
    );
};

export default About;

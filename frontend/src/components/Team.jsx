import React from 'react';
import { motion } from 'framer-motion';
import { Users } from 'lucide-react';
import harshImage from '../assets/harsh-hiremath.jpg';
import zainImage from '../assets/zain-jawad.jpg';
import shreyaImage from '../assets/shreya-kotagiri.jpg';
import kritikaImage from '../assets/kritika.jpg';
import saiVarshiniImage from '../assets/sai-varshini.jpg';
import srujanImage from '../assets/srujan-l.jpg';
import jaganathImage from '../assets/jaganath-reddy.jpg';
import naveedImage from '../assets/naveed-ahmed.jpg';
import manaviImage from '../assets/manavi-ka.png';
import shoaebImage from '../assets/shoaeb-malik.jpg';
import dishaImage from '../assets/disha-naik.jpg';
import trivikramImage from '../assets/trivikram-kalagi.jpg';
import abhijnaImage from '../assets/abhijna-ym.jpg';
import aishwaryaImage from '../assets/aishwarya-gujjar.jpg';
import shrinidhiImage from '../assets/shrinidhi-badiger.jpg';
import aiswaryaSinghImage from '../assets/aiswarya-singh.jpg';
import tarhatImage from '../assets/tarhat-fatima.jpg';
import harshiniImage from '../assets/harshini-gs.jpg';
import akankshaImage from '../assets/akanksha-kinalkar.jpg';
import chethanImage from '../assets/chethan-hs.jpg';
import parnithaImage from '../assets/parnitha-petnikota.jpg';
import jeevithaImage from '../assets/jeevitha-jagadeesh.jpg';
import nehaImage from '../assets/neha-sinchana.jpg';
import adityaImage from '../assets/aditya-kanti.jpg';
import shashankImage from '../assets/shashank-kashyap.jpg';

const Team = () => {
    const coreMembers = [
        { role: 'President', name: 'Mr. Harsh Hiremath', image: harshImage },
        { role: 'Vice President', name: 'Mr. Zain Jawad', image: zainImage },
        { role: 'Secretary', name: 'Ms. Shreya Kotagiri', image: shreyaImage },
        { role: 'Joint Secretary', name: 'Ms. Kritika Mottadi', image: kritikaImage },
        { role: 'Head of Finance', name: 'Ms. Sai Varshini', image: saiVarshiniImage },
        { role: 'Technical Team Head', name: 'Mr. Srujan L', image: srujanImage },
        { role: 'Head of Media and Design', name: 'Ms. Disha D Naik', image: dishaImage },
        { role: 'Head of PR and Marketing', name: 'Ms. Aiswarya Singh', image: aiswaryaSinghImage },
        { role: 'Head of Event Management', name: 'Mr. Chethan H S', image: chethanImage },
    ];

    const eventManagementCoordinators = [
        { name: 'Ms. Parnitha Petnikota', image: parnithaImage },
        { name: 'Ms. Jeevitha Jagadeesh', image: jeevithaImage },
        { name: 'Ms. Neha Sinchana.D', image: nehaImage },
        { name: 'Mr. Aditya Kanti', image: adityaImage },
        { name: 'Mr. Shashank S Kashyap', image: shashankImage },
    ];

    const prMarketingCoordinators = [
        { name: 'Ms. Tarhat Fatima', image: tarhatImage },
        { name: 'Ms. Harshini GS', image: harshiniImage },
        { name: 'Ms. Akanksha Kinalkar', image: akankshaImage },
    ];

    const technicalCoordinators = [
        { name: 'Mr. Jaganath Thimma Reddy', image: jaganathImage },
        { name: 'Mr. Naveed Ahmed', image: naveedImage },
        { name: 'Ms. Manavi K A', image: manaviImage },
        { name: 'Mr. Mohammed Shoaeb Malik', image: shoaebImage },
    ];

    const mediaDesignCoordinators = [
        { name: 'Mr. Trivikram Kalagi', image: trivikramImage },
        { name: 'Ms. Abhijna YM', image: abhijnaImage },
        { name: 'Ms. Aishwarya Gujjar', image: aishwaryaImage },
        { name: 'Ms. Shrinidhi S Badiger', image: shrinidhiImage },
    ];

    return (
        <section className="section-container">
            <h2 className="text-gradient heading-md">
                Core Team
            </h2>
            <div style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fit, minmax(min(250px, 100%), 1fr))',
                gap: '1.5rem',
                maxWidth: '1200px',
                margin: '0 auto',
                marginBottom: '4rem'
            }}>
                {coreMembers.map((member, index) => (
                    <motion.div
                        key={index}
                        className="glass-card neon-border"
                        initial={{ opacity: 0, scale: 0.9 }}
                        whileInView={{ opacity: 1, scale: 1 }}
                        transition={{ duration: 0.3, delay: index * 0.1 }}
                        style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}
                    >
                        <div style={{
                            width: '150px',
                            height: '150px',
                            borderRadius: '50%',
                            background: 'rgba(255,255,255,0.05)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            marginBottom: '1rem',
                            overflow: 'hidden',
                            border: member.image ? '2px solid var(--secondary-purple)' : 'none'
                        }}>
                            {member.image ? (
                                <img
                                    src={member.image}
                                    alt={member.name}
                                    style={{
                                        width: '100%',
                                        height: '100%',
                                        objectFit: 'cover'
                                    }}
                                />
                            ) : (
                                <Users size={40} color="var(--secondary-purple)" />
                            )}
                        </div>
                        <h4 style={{ fontSize: '1.2rem', marginBottom: '0.5rem' }}>{member.name}</h4>
                        <p style={{ color: 'var(--accent-magenta)', fontSize: '0.9rem' }}>{member.role}</p>
                    </motion.div>
                ))}
            </div>

            {/* Technical Team Coordinators - Horizontal Scroll */}
            <div style={{ marginTop: '3rem' }}>
                <h3 className="text-gradient" style={{ fontSize: '1.5rem', marginBottom: '1.5rem' }}>
                    Technical Team Coordinators
                </h3>
                <div style={{
                    overflowX: 'auto',
                    overflowY: 'hidden',
                    scrollBehavior: 'smooth',
                    WebkitOverflowScrolling: 'touch',
                    padding: '1rem 0',
                    margin: '0 -1rem',
                    paddingLeft: '1rem',
                    paddingRight: '1rem'
                }}>
                    <div style={{
                        display: 'flex',
                        gap: '1.5rem',
                        minWidth: 'min-content'
                    }}>
                        {technicalCoordinators.map((coordinator, index) => (
                            <motion.div
                                key={index}
                                className="glass-card neon-border"
                                initial={{ opacity: 0, x: 50 }}
                                whileInView={{ opacity: 1, x: 0 }}
                                transition={{ duration: 0.3, delay: index * 0.1 }}
                                style={{
                                    display: 'flex',
                                    flexDirection: 'column',
                                    alignItems: 'center',
                                    minWidth: 'min(250px, 70vw)',
                                    flexShrink: 0
                                }}
                            >
                                <div style={{
                                    width: '150px',
                                    height: '150px',
                                    borderRadius: '50%',
                                    background: 'rgba(255,255,255,0.05)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    marginBottom: '1rem',
                                    overflow: 'hidden',
                                    border: '2px solid var(--primary-cyan)'
                                }}>
                                    <img
                                        src={coordinator.image}
                                        alt={coordinator.name}
                                        style={{
                                            width: '100%',
                                            height: '100%',
                                            objectFit: 'cover'
                                        }}
                                    />
                                </div>
                                <h4 style={{ fontSize: '1.1rem', marginBottom: '0.5rem' }}>{coordinator.name}</h4>
                                <p style={{ color: 'var(--primary-cyan)', fontSize: '0.85rem' }}>Technical Coordinator</p>
                            </motion.div>
                        ))}
                    </div>
                </div>
            </div>

            {/* Media and Design Coordinators - Horizontal Scroll */}
            <div style={{ marginTop: '3rem' }}>
                <h3 className="text-gradient" style={{ fontSize: '1.5rem', marginBottom: '1.5rem' }}>
                    Media and Design Coordinators
                </h3>
                <div style={{
                    overflowX: 'auto',
                    overflowY: 'hidden',
                    scrollBehavior: 'smooth',
                    WebkitOverflowScrolling: 'touch',
                    padding: '1rem 0',
                    margin: '0 -1rem',
                    paddingLeft: '1rem',
                    paddingRight: '1rem'
                }}>
                    <div style={{
                        display: 'flex',
                        gap: '1.5rem',
                        minWidth: 'min-content'
                    }}>
                        {mediaDesignCoordinators.map((coordinator, index) => (
                            <motion.div
                                key={index}
                                className="glass-card neon-border"
                                initial={{ opacity: 0, x: 50 }}
                                whileInView={{ opacity: 1, x: 0 }}
                                transition={{ duration: 0.3, delay: index * 0.1 }}
                                style={{
                                    display: 'flex',
                                    flexDirection: 'column',
                                    alignItems: 'center',
                                    minWidth: 'min(250px, 70vw)',
                                    flexShrink: 0
                                }}
                            >
                                <div style={{
                                    width: '150px',
                                    height: '150px',
                                    borderRadius: '50%',
                                    background: 'rgba(255,255,255,0.05)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    marginBottom: '1rem',
                                    overflow: 'hidden',
                                    border: '2px solid var(--accent-magenta)'
                                }}>
                                    <img
                                        src={coordinator.image}
                                        alt={coordinator.name}
                                        style={{
                                            width: '100%',
                                            height: '100%',
                                            objectFit: 'cover'
                                        }}
                                    />
                                </div>
                                <h4 style={{ fontSize: '1.1rem', marginBottom: '0.5rem' }}>{coordinator.name}</h4>
                                <p style={{ color: 'var(--accent-magenta)', fontSize: '0.85rem' }}>Media & Design Coordinator</p>
                            </motion.div>
                        ))}
                    </div>
                </div>
            </div>

            {/* PR and Marketing Coordinators - Horizontal Scroll */}
            <div style={{ marginTop: '3rem' }}>
                <h3 className="text-gradient" style={{ fontSize: '1.5rem', marginBottom: '1.5rem' }}>
                    PR and Marketing Coordinators
                </h3>
                <div style={{
                    overflowX: 'auto',
                    overflowY: 'hidden',
                    scrollBehavior: 'smooth',
                    WebkitOverflowScrolling: 'touch',
                    padding: '1rem 0',
                    margin: '0 -1rem',
                    paddingLeft: '1rem',
                    paddingRight: '1rem'
                }}>
                    <div style={{
                        display: 'flex',
                        gap: '1.5rem',
                        minWidth: 'min-content'
                    }}>
                        {prMarketingCoordinators.map((coordinator, index) => (
                            <motion.div
                                key={index}
                                className="glass-card neon-border"
                                initial={{ opacity: 0, x: 50 }}
                                whileInView={{ opacity: 1, x: 0 }}
                                transition={{ duration: 0.3, delay: index * 0.1 }}
                                style={{
                                    display: 'flex',
                                    flexDirection: 'column',
                                    alignItems: 'center',
                                    minWidth: 'min(250px, 70vw)',
                                    flexShrink: 0
                                }}
                            >
                                <div style={{
                                    width: '150px',
                                    height: '150px',
                                    borderRadius: '50%',
                                    background: 'rgba(255,255,255,0.05)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    marginBottom: '1rem',
                                    overflow: 'hidden',
                                    border: '2px solid #ffb300' /* Amber for PR & Marketing */
                                }}>
                                    <img
                                        src={coordinator.image}
                                        alt={coordinator.name}
                                        style={{
                                            width: '100%',
                                            height: '100%',
                                            objectFit: 'cover'
                                        }}
                                    />
                                </div>
                                <h4 style={{ fontSize: '1.1rem', marginBottom: '0.5rem' }}>{coordinator.name}</h4>
                                <p style={{ color: '#ffb300', fontSize: '0.85rem' }}>PR & Marketing Coordinator</p>
                            </motion.div>
                        ))}
                    </div>
                </div>
            </div>

            {/* Event Management Coordinators - Horizontal Scroll */}
            <div style={{ marginTop: '3rem' }}>
                <h3 className="text-gradient" style={{ fontSize: '1.5rem', marginBottom: '1.5rem' }}>
                    Event Management Coordinators
                </h3>
                <div style={{
                    overflowX: 'auto',
                    overflowY: 'hidden',
                    scrollBehavior: 'smooth',
                    WebkitOverflowScrolling: 'touch',
                    padding: '1rem 0',
                    margin: '0 -1rem',
                    paddingLeft: '1rem',
                    paddingRight: '1rem'
                }}>
                    <div style={{
                        display: 'flex',
                        gap: '1.5rem',
                        minWidth: 'min-content'
                    }}>
                        {eventManagementCoordinators.map((coordinator, index) => (
                            <motion.div
                                key={index}
                                className="glass-card neon-border"
                                initial={{ opacity: 0, x: 50 }}
                                whileInView={{ opacity: 1, x: 0 }}
                                transition={{ duration: 0.3, delay: index * 0.1 }}
                                style={{
                                    display: 'flex',
                                    flexDirection: 'column',
                                    alignItems: 'center',
                                    minWidth: 'min(250px, 70vw)',
                                    flexShrink: 0
                                }}
                            >
                                <div style={{
                                    width: '150px',
                                    height: '150px',
                                    borderRadius: '50%',
                                    background: 'rgba(255,255,255,0.05)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    marginBottom: '1rem',
                                    overflow: 'hidden',
                                    border: '2px solid #4ade80' /* Green for Event Management */
                                }}>
                                    {coordinator.image ? (
                                        <img
                                            src={coordinator.image}
                                            alt={coordinator.name}
                                            style={{
                                                width: '100%',
                                                height: '100%',
                                                objectFit: 'cover'
                                            }}
                                        />
                                    ) : (
                                        <Users size={60} color="#4ade80" />
                                    )}
                                </div>
                                <h4 style={{ fontSize: '1.1rem', marginBottom: '0.5rem' }}>{coordinator.name}</h4>
                                <p style={{ color: '#4ade80', fontSize: '0.85rem' }}>Event Management Coordinator</p>
                            </motion.div>
                        ))}
                    </div>
                </div>
            </div>
        </section>
    );
};

export default Team;

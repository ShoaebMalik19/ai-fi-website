import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { useNavigate, useLocation } from 'react-router-dom';
import { Menu, X } from 'lucide-react';

const Navbar = () => {
    const [scrolled, setScrolled] = useState(false);
    const [isOpen, setIsOpen] = useState(false);
    const navigate = useNavigate();
    const location = useLocation();

    useEffect(() => {
        const handleScroll = () => {
            setScrolled(window.scrollY > 50);
        };
        window.addEventListener('scroll', handleScroll);
        return () => window.removeEventListener('scroll', handleScroll);
    }, []);

    const scrollToSection = (id) => {
        setIsOpen(false);
        if (location.pathname !== '/') {
            navigate('/');
            setTimeout(() => {
                const element = document.getElementById(id);
                if (element) element.scrollIntoView({ behavior: 'smooth' });
            }, 100);
        } else {
            const element = document.getElementById(id);
            if (element) element.scrollIntoView({ behavior: 'smooth' });
        }
    };

    const navLinks = [
        { name: 'Home', id: 'home' },
        { name: 'About', id: 'about' },
        { name: 'Events', id: 'events' },
        { name: 'Contact', id: 'contact' },
    ];

    return (
        <nav
            style={{
                position: 'fixed',
                top: 0,
                left: 0,
                right: 0,
                zIndex: 1000,
                padding: '1rem 2rem',
                transition: 'all 0.3s ease',
                background: scrolled ? 'rgba(0, 0, 0, 0.8)' : 'transparent',
                backdropFilter: scrolled ? 'blur(10px)' : 'none',
                borderBottom: scrolled ? '1px solid rgba(255, 255, 255, 0.1)' : 'none',
            }}
        >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', maxWidth: '1200px', margin: '0 auto' }}>
                <h1
                    className="text-gradient"
                    style={{ fontSize: '1.5rem', cursor: 'pointer', margin: 0 }}
                    onClick={() => scrollToSection('home')}
                >
                    AIFi
                </h1>

                {/* Desktop Menu */}
                <div className="desktop-menu" style={{ display: 'flex', gap: '2rem', alignItems: 'center' }}>
                    {navLinks.map((link) => (
                        <button
                            key={link.name}
                            onClick={() => scrollToSection(link.id)}
                            style={{
                                background: 'none',
                                border: 'none',
                                color: '#fff',
                                cursor: 'pointer',
                                fontSize: '1rem',
                                fontWeight: '500',
                                transition: 'color 0.3s ease',
                            }}
                            onMouseEnter={(e) => e.target.style.color = 'var(--primary-cyan)'}
                            onMouseLeave={(e) => e.target.style.color = '#fff'}
                        >
                            {link.name}
                        </button>
                    ))}
                </div>

                {/* Mobile Menu Toggle */}
                <div className="mobile-menu-toggle" style={{ display: 'none' }}>
                    <button onClick={() => setIsOpen(!isOpen)} style={{ background: 'none', border: 'none', color: '#fff' }}>
                        {isOpen ? <X size={24} /> : <Menu size={24} />}
                    </button>
                </div>
            </div>

            {/* Mobile Menu */}
            <style>{`
                @media (max-width: 768px) {
                    .desktop-menu { display: none !important; }
                    .mobile-menu-toggle { display: block !important; }
                }
            `}</style>

            <AnimatePresence>
                {isOpen && (
                    <motion.div
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: 'auto' }}
                        exit={{ opacity: 0, height: 0 }}
                        style={{
                            background: 'rgba(0, 0, 0, 0.95)',
                            backdropFilter: 'blur(10px)',
                            overflow: 'hidden',
                            position: 'absolute',
                            top: '100%',
                            left: 0,
                            right: 0,
                            borderBottom: '1px solid rgba(255,255,255,0.1)'
                        }}
                    >
                        <div style={{ display: 'flex', flexDirection: 'column', padding: '1rem', gap: '1rem' }}>
                            {navLinks.map((link) => (
                                <button
                                    key={link.name}
                                    onClick={() => scrollToSection(link.id)}
                                    style={{
                                        background: 'none',
                                        border: 'none',
                                        color: '#fff',
                                        cursor: 'pointer',
                                        fontSize: '1.1rem',
                                        textAlign: 'left',
                                        padding: '0.5rem'
                                    }}
                                >
                                    {link.name}
                                </button>
                            ))}
                        </div>
                    </motion.div>
                )}
            </AnimatePresence>
        </nav>
    );
};

export default Navbar;

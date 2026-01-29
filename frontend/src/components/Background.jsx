import React, { useEffect, useRef } from 'react';
import { useAnimation } from '../context/AnimationContext';

const Background = () => {
    const canvasRef = useRef(null);
    const { isWarping } = useAnimation();

    useEffect(() => {
        const canvas = canvasRef.current;
        const ctx = canvas.getContext('2d');
        let animationFrameId;

        let width, height;
        let cx, cy;
        let mouse = { x: null, y: null, radius: 150 };

        // Configuration
        const networkParticleCount = 60; // Reduced from 100
        const connectionDistance = 140;
        const starCount = 300; // Background stars

        // Arrays
        const networkParticles = [];
        const stars = [];

        // Background Star Class (3D)
        class Star {
            constructor() {
                this.x = (Math.random() - 0.5) * width * 2; // Spread wider
                this.y = (Math.random() - 0.5) * height * 2;
                this.z = Math.random() * width; // Depth
                this.pz = this.z; // Previous z for trails
                this.size = 0;
            }

            update(isWarping) {
                this.pz = this.z;

                let speed = isWarping ? 80 : 2; // Warp speed vs Idle speed
                this.z -= speed;

                // Reset if passed camera
                if (this.z < 1) {
                    this.z = width;
                    this.pz = this.z;
                    this.x = (Math.random() - 0.5) * width * 2;
                    this.y = (Math.random() - 0.5) * height * 2;
                }
            }

            draw(isWarping) {
                // Perspective projection
                let sx = (this.x / this.z) * width + cx;
                let sy = (this.y / this.z) * height + cy;

                let r = (1 - this.z / width) * 4; // Size based on proximity

                if (sx < 0 || sx > width || sy < 0 || sy > height) return;

                ctx.beginPath();
                ctx.fillStyle = isWarping ? '#00f2ff' : 'white';

                if (isWarping) {
                    // Draw Streak
                    let px = (this.x / this.pz) * width + cx;
                    let py = (this.y / this.pz) * height + cy;

                    ctx.strokeStyle = '#00f2ff';
                    ctx.lineWidth = r;
                    ctx.moveTo(px, py);
                    ctx.lineTo(sx, sy);
                    ctx.stroke();
                } else {
                    // Draw Dot
                    ctx.arc(sx, sy, r, 0, Math.PI * 2);
                    ctx.fill();
                }
            }
        }

        // Interactive Network Particle Class
        class NetworkParticle {
            constructor() {
                this.x = Math.random() * width;
                this.y = Math.random() * height;
                this.vx = (Math.random() - 0.5) * 0.8;
                this.vy = (Math.random() - 0.5) * 0.8;
                this.size = Math.random() * 2 + 1.5;
                this.color = '#00f2ff'; // Cyan
                this.opacity = 1;
            }

            update(isWarping) {
                if (isWarping) {
                    this.opacity -= 0.1;
                    if (this.opacity < 0) this.opacity = 0;
                } else {
                    this.opacity += 0.1;
                    if (this.opacity > 1) this.opacity = 1;

                    this.x += this.vx;
                    this.y += this.vy;
                }

                if (this.x < 0 || this.x > width) this.vx *= -1;
                if (this.y < 0 || this.y > height) this.vy *= -1;

                // Mouse interaction
                if (mouse.x != null) {
                    let dx = mouse.x - this.x;
                    let dy = mouse.y - this.y;
                    let distance = Math.sqrt(dx * dx + dy * dy);

                    if (distance < mouse.radius) {
                        const forceDirectionX = dx / distance;
                        const forceDirectionY = dy / distance;
                        const force = (mouse.radius - distance) / mouse.radius;
                        // Push away gentler
                        const directionX = forceDirectionX * force * 2;
                        const directionY = forceDirectionY * force * 2;

                        this.x -= directionX;
                        this.y -= directionY;
                    }
                }
            }

            draw() {
                if (this.opacity <= 0) return;
                ctx.beginPath();
                ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
                ctx.fillStyle = `rgba(0, 242, 255, ${this.opacity})`;
                ctx.fill();
            }
        }

        const init = () => {
            width = window.innerWidth;
            height = window.innerHeight;
            cx = width / 2;
            cy = height / 2;
            canvas.width = width;
            canvas.height = height;

            networkParticles.length = 0;
            stars.length = 0;

            for (let i = 0; i < networkParticleCount; i++) {
                networkParticles.push(new NetworkParticle());
            }
            for (let i = 0; i < starCount; i++) {
                stars.push(new Star());
            }
        };

        const drawMilkyWay = () => {
            // Draw a diagonal cloudy band
            ctx.fillStyle = 'rgba(0, 0, 0, 0.2)'; // Clear accumulation

            const gradient = ctx.createLinearGradient(0, 0, width, height);
            gradient.addColorStop(0, 'rgba(15, 12, 41, 1)');
            gradient.addColorStop(0.3, 'rgba(48, 43, 99, 0.4)'); // Purple haze
            gradient.addColorStop(0.5, 'rgba(36, 36, 62, 0.2)');
            gradient.addColorStop(0.7, 'rgba(48, 43, 99, 0.4)');
            gradient.addColorStop(1, 'rgba(15, 12, 41, 1)');

            ctx.fillStyle = gradient;
            ctx.fillRect(0, 0, width, height);

            // Add some "Nebula" clouds
            ctx.filter = 'blur(60px)';
            ctx.fillStyle = 'rgba(126, 34, 206, 0.15)'; // Purple
            ctx.beginPath();
            ctx.arc(width * 0.2, height * 0.3, 300, 0, Math.PI * 2);
            ctx.fill();

            ctx.fillStyle = 'rgba(0, 242, 255, 0.1)'; // Cyan
            ctx.beginPath();
            ctx.arc(width * 0.8, height * 0.7, 400, 0, Math.PI * 2);
            ctx.fill();

            ctx.filter = 'none'; // Reset filter
        };

        const connectParticles = () => {
            for (let a = 0; a < networkParticles.length; a++) {
                for (let b = a; b < networkParticles.length; b++) {
                    let dx = networkParticles[a].x - networkParticles[b].x;
                    let dy = networkParticles[a].y - networkParticles[b].y;
                    let distance = Math.sqrt(dx * dx + dy * dy);

                    if (distance < connectionDistance) {
                        ctx.strokeStyle = `rgba(0, 242, 255, ${1 - distance / connectionDistance})`;
                        ctx.lineWidth = 0.8;
                        ctx.beginPath();
                        ctx.moveTo(networkParticles[a].x, networkParticles[a].y);
                        ctx.lineTo(networkParticles[b].x, networkParticles[b].y);
                        ctx.stroke();
                    }
                }
                // Mouse
                if (mouse.x != null) {
                    let dx = networkParticles[a].x - mouse.x;
                    let dy = networkParticles[a].y - mouse.y;
                    let distance = Math.sqrt(dx * dx + dy * dy);

                    if (distance < mouse.radius) {
                        ctx.strokeStyle = `rgba(255, 0, 204, ${1 - distance / mouse.radius})`;
                        ctx.lineWidth = 1;
                        ctx.beginPath();
                        ctx.moveTo(networkParticles[a].x, networkParticles[a].y);
                        ctx.lineTo(mouse.x, mouse.y);
                        ctx.stroke();
                    }
                }
            }
        };

        const animate = () => {
            ctx.clearRect(0, 0, width, height); // Clear

            // 1. Draw Base (Milky Way / Nebula)
            drawMilkyWay();

            // 2. Draw Background Stars (Hyperspace)
            stars.forEach(star => {
                star.update(isWarping);
                star.draw(isWarping);
            });

            // 3. Draw Network
            networkParticles.forEach(particle => {
                particle.update(isWarping);
                particle.draw();
            });

            connectParticles();
            animationFrameId = requestAnimationFrame(animate);
        };

        init();
        animate();

        const handleResize = () => init();
        const handleMouseMove = (e) => {
            mouse.x = e.x;
            mouse.y = e.y;
        };
        const handleMouseLeave = () => {
            mouse.x = null;
            mouse.y = null;
        }

        window.addEventListener('resize', handleResize);
        window.addEventListener('mousemove', handleMouseMove);
        window.addEventListener('mouseleave', handleMouseLeave);

        return () => {
            window.removeEventListener('resize', handleResize);
            window.removeEventListener('mousemove', handleMouseMove);
            window.removeEventListener('mouseleave', handleMouseLeave);
            cancelAnimationFrame(animationFrameId);
        };
    }, [isWarping]);

    return (
        <canvas
            ref={canvasRef}
            style={{
                position: 'fixed',
                top: 0,
                left: 0,
                zIndex: -1,
                width: '100vw',
                height: '100vh',
                pointerEvents: 'none'
            }}
        />
    );
};

export default Background;

// SB Consulting - Premium Organization Diagnosis Website
// Enhanced JavaScript with Animations

document.addEventListener('DOMContentLoaded', function() {

    // ==================== DROPDOWN MENU ====================
    document.querySelectorAll('.nav-links li.has-dropdown').forEach(function(item) {
        var menu = item.querySelector('.dropdown-menu');
        if (!menu) return;

        function showMenu() {
            var rect = item.getBoundingClientRect();
            menu.style.top  = (rect.bottom + 4) + 'px';
            menu.style.left = rect.left + 'px';
            menu.classList.add('show');
        }
        function hideMenu() {
            menu.classList.remove('show');
        }

        item.addEventListener('mouseenter', showMenu);
        item.addEventListener('mouseleave', function(e) {
            if (!menu.contains(e.relatedTarget)) hideMenu();
        });
        menu.addEventListener('mouseleave', function(e) {
            if (!item.contains(e.relatedTarget)) hideMenu();
        });
        menu.addEventListener('mouseenter', function() {
            menu.classList.add('show');
        });
    });

    // ==================== MOBILE MENU ====================
    const mobileMenuBtn = document.querySelector('.mobile-menu-btn');
    const navLinks = document.querySelector('.nav-links');

    if (mobileMenuBtn) {
        mobileMenuBtn.addEventListener('click', function() {
            navLinks.classList.toggle('active');
            this.classList.toggle('active');

            // Animate hamburger menu
            const spans = this.querySelectorAll('span');
            if (this.classList.contains('active')) {
                spans[0].style.transform = 'rotate(45deg) translate(6px, 6px)';
                spans[1].style.opacity = '0';
                spans[2].style.transform = 'rotate(-45deg) translate(6px, -6px)';
            } else {
                spans[0].style.transform = 'none';
                spans[1].style.opacity = '1';
                spans[2].style.transform = 'none';
            }
        });
    }

    // Close mobile menu when clicking outside
    document.addEventListener('click', function(e) {
        if (!e.target.closest('nav') && navLinks) {
            navLinks.classList.remove('active');
            if (mobileMenuBtn) {
                mobileMenuBtn.classList.remove('active');
                const spans = mobileMenuBtn.querySelectorAll('span');
                spans[0].style.transform = 'none';
                spans[1].style.opacity = '1';
                spans[2].style.transform = 'none';
            }
        }
    });

    // ==================== HEADER SCROLL EFFECT ====================
    const header = document.querySelector('header');
    let lastScroll = 0;

    window.addEventListener('scroll', function() {
        const currentScroll = window.pageYOffset;

        if (currentScroll > 50) {
            header.classList.add('scrolled');
        } else {
            header.classList.remove('scrolled');
        }

        lastScroll = currentScroll;
    });

    // ==================== SMOOTH SCROLL ====================
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function(e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                const headerHeight = header.offsetHeight;
                const targetPosition = target.getBoundingClientRect().top + window.pageYOffset - headerHeight;

                window.scrollTo({
                    top: targetPosition,
                    behavior: 'smooth'
                });
            }
        });
    });

    // ==================== INTERSECTION OBSERVER FOR ANIMATIONS ====================
    const observerOptions = {
        root: null,
        rootMargin: '0px',
        threshold: 0.1
    };

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('visible');
            }
        });
    }, observerOptions);

    // Observe all animated elements
    document.querySelectorAll('.fade-in, .slide-in-left, .slide-in-right').forEach(el => {
        observer.observe(el);
    });

    // Add animation to cards with stagger effect
    const animatedElements = document.querySelectorAll('.feature-card, .area-card, .outcome-card, .roadmap-item, .model-item');
    animatedElements.forEach((el, index) => {
        el.classList.add('fade-in');
        el.style.transitionDelay = `${index * 0.08}s`;
        observer.observe(el);
    });

    // ==================== ACTIVE NAVIGATION LINK ====================
    const currentPage = window.location.pathname.split('/').pop() || 'index.html';
    document.querySelectorAll('.nav-links a').forEach(link => {
        const href = link.getAttribute('href');
        const linkPage = href.split('/').pop();

        if (linkPage === currentPage ||
            (currentPage === '' && linkPage === 'index.html') ||
            (currentPage === 'index.html' && href === 'index.html')) {
            link.classList.add('active');
        }
    });

    // ==================== FORM HANDLING ====================
    const contactForm = document.querySelector('.contact-form form');
    if (contactForm) {
        contactForm.addEventListener('submit', function(e) {
            e.preventDefault();

            // Get form data
            const formData = new FormData(this);
            const data = {};
            formData.forEach((value, key) => {
                data[key] = value;
            });

            // Show loading state
            const submitBtn = this.querySelector('.submit-btn');
            const originalText = submitBtn.textContent;
            submitBtn.textContent = '전송 중...';
            submitBtn.disabled = true;

            // Simulate form submission
            setTimeout(() => {
                // Show success message
                showNotification('문의가 성공적으로 접수되었습니다. 담당자가 빠른 시일 내에 연락드리겠습니다.', 'success');
                this.reset();
                submitBtn.textContent = originalText;
                submitBtn.disabled = false;
            }, 1500);
        });
    }

    // ==================== NOTIFICATION SYSTEM ====================
    function showNotification(message, type = 'info') {
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <div class="notification-content">
                <span class="notification-icon">${type === 'success' ? '&#10003;' : '&#8505;'}</span>
                <span class="notification-message">${message}</span>
            </div>
        `;

        notification.style.cssText = `
            position: fixed;
            top: 100px;
            right: 30px;
            background: ${type === 'success' ? 'linear-gradient(135deg, #059669 0%, #047857 100%)' : 'linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)'};
            color: white;
            padding: 20px 28px;
            border-radius: 16px;
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.2);
            z-index: 10000;
            transform: translateX(150%);
            transition: transform 0.5s cubic-bezier(0.4, 0, 0.2, 1);
            max-width: 400px;
            font-weight: 500;
        `;

        document.body.appendChild(notification);

        // Animate in
        setTimeout(() => {
            notification.style.transform = 'translateX(0)';
        }, 100);

        // Animate out and remove
        setTimeout(() => {
            notification.style.transform = 'translateX(150%)';
            setTimeout(() => {
                notification.remove();
            }, 500);
        }, 4000);
    }

    // ==================== COUNTER ANIMATION ====================
    function animateCounter(element, target, duration = 2000, suffix = '') {
        let start = 0;
        const increment = target / (duration / 16);
        const startTime = performance.now();

        function updateCounter(currentTime) {
            const elapsed = currentTime - startTime;
            const progress = Math.min(elapsed / duration, 1);
            const easeOutQuart = 1 - Math.pow(1 - progress, 4);
            const currentValue = Math.floor(easeOutQuart * target);

            element.textContent = currentValue + suffix;

            if (progress < 1) {
                requestAnimationFrame(updateCounter);
            } else {
                element.textContent = target + suffix;
            }
        }

        requestAnimationFrame(updateCounter);
    }

    // Trigger counter animation when visible
    const counters = document.querySelectorAll('.counter, .stat-number');
    counters.forEach(counter => {
        const target = parseInt(counter.getAttribute('data-target')) || parseInt(counter.textContent);
        const suffix = counter.getAttribute('data-suffix') || '';

        const counterObserver = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    animateCounter(counter, target, 2000, suffix);
                    counterObserver.unobserve(entry.target);
                }
            });
        }, { threshold: 0.5 });

        counterObserver.observe(counter);
    });

    // ==================== SCROLL TO TOP BUTTON ====================
    const scrollTopBtn = document.createElement('button');
    scrollTopBtn.innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="18,15 12,9 6,15"></polyline></svg>';
    scrollTopBtn.className = 'scroll-top-btn';
    document.body.appendChild(scrollTopBtn);

    scrollTopBtn.addEventListener('click', () => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    window.addEventListener('scroll', () => {
        if (window.pageYOffset > 500) {
            scrollTopBtn.classList.add('visible');
        } else {
            scrollTopBtn.classList.remove('visible');
        }
    });

    // ==================== PARALLAX EFFECT FOR HERO ====================
    const hero = document.querySelector('.hero');
    if (hero) {
        window.addEventListener('scroll', () => {
            const scrolled = window.pageYOffset;
            const heroContent = hero.querySelector('.hero-content');
            if (heroContent && scrolled < hero.offsetHeight) {
                heroContent.style.transform = `translateY(${scrolled * 0.3}px)`;
                heroContent.style.opacity = 1 - (scrolled / hero.offsetHeight) * 0.5;
            }
        });
    }

    // ==================== TYPING EFFECT ====================
    function typeWriter(element, text, speed = 50) {
        let i = 0;
        element.textContent = '';

        function type() {
            if (i < text.length) {
                element.textContent += text.charAt(i);
                i++;
                setTimeout(type, speed);
            }
        }

        type();
    }

    // ==================== HOVER EFFECTS FOR CARDS ====================
    document.querySelectorAll('.feature-card, .outcome-card, .area-card').forEach(card => {
        card.addEventListener('mouseenter', function(e) {
            this.style.transform = 'translateY(-10px) scale(1.02)';
        });

        card.addEventListener('mouseleave', function(e) {
            this.style.transform = '';
        });
    });

    // ==================== IMAGE LAZY LOADING ====================
    const lazyImages = document.querySelectorAll('img[data-src]');
    const imageObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const img = entry.target;
                img.src = img.dataset.src;
                img.removeAttribute('data-src');
                imageObserver.unobserve(img);
            }
        });
    });

    lazyImages.forEach(img => imageObserver.observe(img));

    // ==================== TOOLTIP SYSTEM ====================
    document.querySelectorAll('[data-tooltip]').forEach(element => {
        element.addEventListener('mouseenter', function(e) {
            const tooltip = document.createElement('div');
            tooltip.className = 'tooltip';
            tooltip.textContent = this.getAttribute('data-tooltip');
            tooltip.style.cssText = `
                position: absolute;
                background: #1e293b;
                color: white;
                padding: 8px 16px;
                border-radius: 8px;
                font-size: 0.85rem;
                white-space: nowrap;
                z-index: 10000;
                opacity: 0;
                transition: opacity 0.3s ease;
                pointer-events: none;
            `;
            document.body.appendChild(tooltip);

            const rect = this.getBoundingClientRect();
            tooltip.style.top = `${rect.top - tooltip.offsetHeight - 10 + window.pageYOffset}px`;
            tooltip.style.left = `${rect.left + (rect.width / 2) - (tooltip.offsetWidth / 2)}px`;

            setTimeout(() => {
                tooltip.style.opacity = '1';
            }, 10);

            this.addEventListener('mouseleave', function() {
                tooltip.remove();
            }, { once: true });
        });
    });

    // ==================== PROGRESS BAR ON SCROLL ====================
    const progressBar = document.createElement('div');
    progressBar.className = 'scroll-progress';
    progressBar.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        height: 3px;
        background: linear-gradient(90deg, #2563eb 0%, #3b82f6 100%);
        z-index: 10001;
        transition: width 0.1s ease;
        width: 0%;
    `;
    document.body.appendChild(progressBar);

    window.addEventListener('scroll', () => {
        const scrollHeight = document.documentElement.scrollHeight - window.innerHeight;
        const scrolled = (window.pageYOffset / scrollHeight) * 100;
        progressBar.style.width = `${scrolled}%`;
    });

    console.log('%c SB Consulting Website Loaded Successfully! ', 'background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); color: white; padding: 10px 20px; border-radius: 8px; font-weight: bold;');
});

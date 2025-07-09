/**
 * ExcelRefinery Styling Template JavaScript
 * File: styling-template.js
 * Description: Interactive functionality for styling template components
 * Author: ExcelRefinery Development Team
 */

(function() {
    'use strict';
    
    // Cache DOM elements
    let cachedElements = {};
    
    // Private methods
    const initializeElements = function() {
        try {
            cachedElements = {
                accordionHeaders: document.querySelectorAll('.accordion-excel-header'),
                navItems: document.querySelectorAll('.nav-excel-item'),
                progressBars: document.querySelectorAll('.progress-excel-bar'),
                alertCloseButtons: document.querySelectorAll('.alert-close'),
                buttons: document.querySelectorAll('.btn-excel'),
                cards: document.querySelectorAll('.card-excel')
            };
        } catch (error) {
            console.error('Error initializing DOM elements:', error);
        }
    };
    
    const initializeAccordion = function() {
        try {
            cachedElements.accordionHeaders.forEach(function(header) {
                header.addEventListener('click', function(e) {
                    e.preventDefault();
                    
                    const content = this.nextElementSibling;
                    const icon = this.querySelector('.accordion-icon');
                    
                    // Toggle active state
                    this.classList.toggle('active');
                    content.classList.toggle('active');
                    
                    // Update icon
                    if (icon) {
                        icon.textContent = content.classList.contains('active') ? 'expand_less' : 'expand_more';
                    }
                    
                    // Close other accordion items (optional - for single-open behavior)
                    // Comment out these lines if you want multiple accordions open at once
                    cachedElements.accordionHeaders.forEach(function(otherHeader) {
                        if (otherHeader !== header) {
                            otherHeader.classList.remove('active');
                            const otherContent = otherHeader.nextElementSibling;
                            const otherIcon = otherHeader.querySelector('.accordion-icon');
                            if (otherContent) {
                                otherContent.classList.remove('active');
                            }
                            if (otherIcon) {
                                otherIcon.textContent = 'expand_more';
                            }
                        }
                    });
                });
            });
        } catch (error) {
            console.error('Error initializing accordion:', error);
        }
    };
    
    const initializeNavigation = function() {
        try {
            cachedElements.navItems.forEach(function(item) {
                item.addEventListener('click', function(e) {
                    e.preventDefault();
                    
                    // Remove active class from all items
                    cachedElements.navItems.forEach(function(navItem) {
                        navItem.classList.remove('active');
                    });
                    
                    // Add active class to clicked item
                    this.classList.add('active');
                });
            });
        } catch (error) {
            console.error('Error initializing navigation:', error);
        }
    };
    
    const animateProgressBars = function() {
        try {
            cachedElements.progressBars.forEach(function(bar, index) {
                // Animate progress bars with different values for demonstration
                const progressValues = [25, 50, 75, 90, 100];
                const targetWidth = progressValues[index % progressValues.length] || 50;
                
                setTimeout(function() {
                    bar.style.width = targetWidth + '%';
                }, 100 * (index + 1));
            });
        } catch (error) {
            console.error('Error animating progress bars:', error);
        }
    };
    
    const initializeInteractiveElements = function() {
        try {
            // Add ripple effect to buttons
            cachedElements.buttons.forEach(function(button) {
                button.addEventListener('click', function(e) {
                    const ripple = document.createElement('span');
                    const rect = this.getBoundingClientRect();
                    const size = Math.max(rect.width, rect.height);
                    const x = e.clientX - rect.left - size / 2;
                    const y = e.clientY - rect.top - size / 2;
                    
                    ripple.style.width = ripple.style.height = size + 'px';
                    ripple.style.left = x + 'px';
                    ripple.style.top = y + 'px';
                    ripple.classList.add('ripple');
                    
                    this.appendChild(ripple);
                    
                    setTimeout(function() {
                        ripple.remove();
                    }, 600);
                });
            });
            
            // Add hover effects to cards
            cachedElements.cards.forEach(function(card) {
                card.addEventListener('mouseenter', function() {
                    this.style.opacity = '0.95';
                });
                
                card.addEventListener('mouseleave', function() {
                    this.style.opacity = '';
                });
            });
            
        } catch (error) {
            console.error('Error initializing interactive elements:', error);
        }
    };
    
    const copyColorCode = function(colorCode, element) {
        try {
            navigator.clipboard.writeText(colorCode).then(function() {
                const originalText = element.textContent;
                element.textContent = 'Copied!';
                element.style.color = '#4caf50';
                
                setTimeout(function() {
                    element.textContent = originalText;
                    element.style.color = '';
                }, 1000);
            }).catch(function(error) {
                console.error('Failed to copy color code:', error);
                // Fallback for older browsers
                const textArea = document.createElement('textarea');
                textArea.value = colorCode;
                document.body.appendChild(textArea);
                textArea.focus();
                textArea.select();
                try {
                    document.execCommand('copy');
                    const originalText = element.textContent;
                    element.textContent = 'Copied!';
                    element.style.color = '#4caf50';
                    
                    setTimeout(function() {
                        element.textContent = originalText;
                        element.style.color = '';
                    }, 1000);
                } catch (fallbackError) {
                    console.error('Fallback copy failed:', fallbackError);
                }
                document.body.removeChild(textArea);
            });
        } catch (error) {
            console.error('Error copying color code:', error);
        }
    };
    
    const initializeColorSwatches = function() {
        try {
            const colorSwatches = document.querySelectorAll('.color-swatch');
            colorSwatches.forEach(function(swatch) {
                swatch.addEventListener('click', function() {
                    const colorInfo = this.querySelector('.color-info');
                    const colorCode = colorInfo.dataset.color;
                    if (colorCode) {
                        copyColorCode(colorCode, colorInfo);
                    }
                });
            });
        } catch (error) {
            console.error('Error initializing color swatches:', error);
        }
    };
    
    // Public methods
    const StylingTemplate = {
        init: function() {
            try {
                initializeElements();
                initializeAccordion();
                initializeNavigation();
                initializeInteractiveElements();
                initializeColorSwatches();
                
                // Animate progress bars after a short delay
                setTimeout(animateProgressBars, 500);
                
                console.log('Styling Template initialized successfully');
            } catch (error) {
                console.error('Error initializing Styling Template:', error);
            }
        },
        
        // Method to refresh components if needed
        refresh: function() {
            try {
                initializeElements();
                console.log('Styling Template refreshed');
            } catch (error) {
                console.error('Error refreshing Styling Template:', error);
            }
        },
        
        // Method to show a demonstration alert
        showDemoAlert: function(type, message) {
            try {
                const alertContainer = document.querySelector('.demo-alerts');
                if (!alertContainer) return;
                
                const alert = document.createElement('div');
                alert.className = `alert-excel alert-excel-${type}`;
                alert.innerHTML = `
                    <i class="material-icons">${this.getAlertIcon(type)}</i>
                    <span>${message}</span>
                `;
                
                alertContainer.appendChild(alert);
                
                // Auto-remove after 3 seconds
                setTimeout(function() {
                    alert.style.opacity = '0';
                    alert.style.transform = 'translateX(100%)';
                    setTimeout(function() {
                        if (alert.parentNode) {
                            alert.parentNode.removeChild(alert);
                        }
                    }, 300);
                }, 3000);
            } catch (error) {
                console.error('Error showing demo alert:', error);
            }
        },
        
        getAlertIcon: function(type) {
            const icons = {
                'success': 'check_circle',
                'warning': 'warning',
                'danger': 'error',
                'info': 'info'
            };
            return icons[type] || 'info';
        }
    };
    
    // Auto-initialize when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', StylingTemplate.init);
    } else {
        StylingTemplate.init();
    }
    
    // Make StylingTemplate available globally for manual calls if needed
    window.StylingTemplate = StylingTemplate;
    
})(); 
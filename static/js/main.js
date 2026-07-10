/**
 * ARCHIVO: main.js
 * FUNCIÓN: Controlador lógico del lado del cliente (Frontend). Gestiona la 
 *          interactividad y la experiencia de usuario (UX) mediante JavaScript.
 * 
 * MÓDULOS PRINCIPALES:
 *  1. Gestor de Temas: Alterna entre Modo Oscuro y Claro, persistiendo la preferencia en localStorage.
 *  2. Buscador en Vivo (Live Search): Filtra tablas de datos en tiempo real sin recargar la página.
 *  3. Notificaciones Toast: Crea alertas flotantes dinámicas (Feedback visual) basándose en los mensajes flash del servidor (Flask).
 *  4. DataTables: Inicializa automáticamente la paginación, ordenamiento y búsqueda profesional para todas las tablas del sistema.
 */

document.addEventListener('DOMContentLoaded', () => {
  // 1. MODO OSCURO INTELIGENTE: cambia la apariencia visual y guarda la preferencia en el navegador.
  const themeToggle = document.getElementById('theme-toggle');
  const body = document.documentElement;

  // Revisar memoria del navegador
  if (localStorage.getItem('theme') === 'dark') {
    body.setAttribute('data-theme', 'dark');
    if (themeToggle) themeToggle.innerHTML = '<i class="bi bi-sun-fill text-warning"></i>';
  }

  // Alternador
  if (themeToggle) {
    themeToggle.addEventListener('click', () => {
      if (body.getAttribute('data-theme') === 'dark') {
        body.removeAttribute('data-theme');
        localStorage.setItem('theme', 'light');
        themeToggle.innerHTML = '<i class="bi bi-moon-stars-fill"></i>';
      } else {
        body.setAttribute('data-theme', 'dark');
        localStorage.setItem('theme', 'dark');
        themeToggle.innerHTML = '<i class="bi bi-sun-fill text-warning"></i>';
      }
    });
  }

  // 2. BUSCADOR EN VIVO (Tablas)
  const searchInputs = document.querySelectorAll('.live-search');
  searchInputs.forEach(input => {
    input.addEventListener('keyup', function() {
      const filter = this.value.toLowerCase();
      const table = document.querySelector(this.dataset.table);
      if (!table) return;
      
      const rows = table.querySelectorAll('tbody tr');
      rows.forEach(row => {
        const text = row.textContent.toLowerCase();
        row.style.display = text.includes(filter) ? '' : 'none';
      });
    });
  });

  // 3. GENERADOR DE NOTIFICACIONES TOAST (Reparado)
  const flashMessages = document.querySelectorAll('.flash-message-hidden');
  const toastContainer = document.getElementById('toast-container');
  
  if (flashMessages.length > 0 && toastContainer) {
    flashMessages.forEach((msg, index) => {
      setTimeout(() => {
        // Usamos getAttribute para evitar conflictos con caracteres especiales
        const text = msg.getAttribute('data-message');
        const category = msg.getAttribute('data-category');
        
        if (!text) return; // Si está vacío, lo ignoramos

        const toast = document.createElement('div');
        let borderColor = category === 'success' ? '#198754' : (category === 'danger' ? '#dc3545' : '#ffc107');
        toast.style.borderLeft = `5px solid ${borderColor}`;
        toast.className = 'custom-toast shadow-sm';
        
        let icon = 'bi-info-circle-fill text-primary';
        if(category === 'success') icon = 'bi-check-circle-fill text-success';
        if(category === 'danger') icon = 'bi-x-circle-fill text-danger';
        if(category === 'warning') icon = 'bi-exclamation-triangle-fill text-warning';

        toast.innerHTML = `<i class="bi ${icon} fs-4"></i> <div>${text}</div>`;
        toastContainer.appendChild(toast);

        // Desaparece a los 4 segundos
        setTimeout(() => {
          toast.classList.add('toast-fade-out');
          setTimeout(() => toast.remove(), 500);
        }, 4000);
      }, index * 200);
    });
  }

  // 4. PROGRESS BARS DEL DASHBOARD
  const progressBars = document.querySelectorAll('.progress-bar[data-value]');
  progressBars.forEach((bar) => {
    const value = Number(bar.dataset.value) || 0;
    bar.style.width = `${value}%`;
    bar.setAttribute('aria-valuenow', value);
  });
});
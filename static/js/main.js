// Toggle sidebar en móviles
document.addEventListener('DOMContentLoaded', function() {
    const toggleBtn = document.getElementById('menu-toggle');
    const sidebar = document.getElementById('sidebar');
    const closeSidebar = document.getElementById('close-sidebar');

    if (toggleBtn) {
        toggleBtn.addEventListener('click', () => {
            sidebar.classList.toggle('open');
        });
    }
    if (closeSidebar) {
        closeSidebar.addEventListener('click', () => {
            sidebar.classList.remove('open');
        });
    }
});

// Confirmaciones
function confirmarEliminar(mensaje) {
    return confirm(mensaje || '¿Estás seguro?');
}

console.log('Texto buscado:', filtro);
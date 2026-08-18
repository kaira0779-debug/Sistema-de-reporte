// particles.js
document.addEventListener('DOMContentLoaded', () => {
    const canvas = document.getElementById('particles-canvas');
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    let particlesArray = [];
    const numberOfParticles = 100;

    // Ajustar tamaño del canvas
    function resizeCanvas() {
        canvas.width = window.innerWidth;
        canvas.height = window.innerHeight;
    }
    window.addEventListener('resize', resizeCanvas);
    resizeCanvas();

    // Clase Partícula
    class Particle {
        constructor() {
            this.x = Math.random() * canvas.width;
            this.y = Math.random() * canvas.height;
            this.size = Math.random() * 3 + 1;
            this.speedX = Math.random() * 1 - 0.5;
            this.speedY = Math.random() * 1 - 0.5;
            this.color = 'rgba(241, 196, 15,' + (Math.random() * 0.5 + 0.2) + ')'; // amarillo semitransparente
        }
        update() {
            this.x += this.speedX;
            this.y += this.speedY;
            if (this.x > canvas.width + 5) this.x = -5;
            if (this.x < -5) this.x = canvas.width + 5;
            if (this.y > canvas.height + 5) this.y = -5;
            if (this.y < -5) this.y = canvas.height + 5;
        }
        draw() {
            ctx.fillStyle = this.color;
            ctx.beginPath();
            ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
            ctx.fill();
        }
    }

    function init() {
        particlesArray = [];
        for (let i = 0; i < numberOfParticles; i++) {
            particlesArray.push(new Particle());
        }
    }

    function animate() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        for (let particle of particlesArray) {
            particle.update();
            particle.draw();
        }
        requestAnimationFrame(animate);
    }

    init();
    animate();

    console.log('Texto buscado:', filtro);

    function seleccionarCliente(cliente) {
    $clienteId.val(cliente.id);
    $inputCliente.val(cliente.nombre); // Asegúrate de que 'cliente.nombre' existe
    $('#cliente_nombre_busqueda').val(cliente.nombre); // Respaldo
    actualizarNodo(cliente.nodo || '');
    $resultadosClientes.hide().empty();
    $mensajeCliente.text('').hide();
    mostrarCliente(cliente);
    console.log('Cliente seleccionado:', cliente);
}
});
# Zona 96K — Sistema de Inventario y Punto de Venta

Sistema de inventario y punto de venta (POS) construido para un negocio real tipo bar/licorería, desarrollado con Google Apps Script y Google Sheets como base de datos.

## 🎯 Qué resuelve

El negocio necesitaba una forma de gestionar ventas, inventario y mesas sin depender de software de POS costoso. Este sistema convierte Google Sheets en una base de datos funcional, con una interfaz web personalizada para el día a día del negocio.

## ⚙️ Funcionalidades principales

- **Venta basada en carrito**: soporte para múltiples productos por transacción, con búsqueda mediante dropdowns dinámicos.
- **Pagos divididos (split payment)**: una misma venta puede pagarse combinando varios métodos de pago (efectivo, tarjeta, transferencia).
- **Gestión de mesas**: seguimiento de mesas activas para el flujo de atención tipo bar/restaurante.
- **Reportería con KPIs**: dashboard con indicadores clave filtrables por rango de fechas.

## 🛠️ Stack técnico

- **Backend**: Google Apps Script
- **Base de datos**: Google Sheets
- **Frontend**: HTML/CSS/JavaScript (interfaz personalizada, sin frameworks)
- **Testing**: arnés de pruebas en Node.js con entorno mock de Google Apps Script — 257 pruebas automatizadas cubriendo la lógica de negocio

## 🧩 Decisiones de diseño

El sistema pasó por una revisión estructural importante: el diseño original agrupaba productos en celdas combinadas dentro de la hoja de cálculo, lo que generaba inestabilidad al escalar. Se migró a un modelo de una fila por producto, lo que simplificó las consultas y eliminó errores intermitentes de sincronización.

## 📌 Estado

Proyecto funcional en uso por el negocio. Desarrollo y mantenimiento continuo.

<img src="images/Alertas.jpeg" width="300" alt="Alerta del sistema">
<img src="images/Inicio.jpeg" width="300" alt="Inicio del sistema">

---

*Desarrollado por [Iván Castro Ujueta](https://github.com/Ivancastr8)*

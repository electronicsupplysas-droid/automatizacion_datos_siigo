# Plan de implementación

---

## Lo que necesito del cliente

| # | Qué necesito | Para qué |
|---|---|---|
| 1 | Un correo corporativo de la empresa | Crear todas las cuentas bajo su propiedad |
| 2 | Acceso a Siigo con perfil administrador | Generar la llave de conexión a la API |
| 3 | Confirmar que el plan de Siigo tiene API habilitada | Sin esto no podemos leer los datos |
| 4 | Aprobar la creación de cuenta en GitHub con ese correo | Guardar el código bajo su repositorio |
| 5 | Aprobar la creación de cuenta en Supabase con ese correo | Base de datos bajo su propiedad |
| 6 | Acceso temporal como colaborador una vez creadas las cuentas | Para que yo configure todo |

---

## Lo que yo hago

1. Creo y configuro el repositorio de código bajo la cuenta GitHub del cliente.
2. Creo y configuro la base de datos bajo la cuenta Supabase del cliente.
3. Cargo el histórico de Siigo en la base de datos.
4. Activo la sincronización automática diaria.
5. Conecto los reportes al dashboard.
6. Entrego todo documentado y funcionando.

---

> Una vez terminada la implementación el cliente tiene el control total.  
> Pueden retirarme el acceso en cualquier momento y el sistema sigue funcionando solo.

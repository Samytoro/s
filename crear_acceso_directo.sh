#!/bin/bash
# Script para crear acceso directo en el escritorio

DESKTOP="$HOME/Desktop"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Crear script ejecutable en el escritorio
cat > "$DESKTOP/Merge F42.command" << 'EOF'
#!/bin/bash
# Obtener directorio donde está este script
cd "$(dirname "$0")"

# Si gui_merge.py no está aquí, buscar en la ubicación original
if [ ! -f "gui_merge.py" ]; then
    # Buscar el archivo en el sistema
    GUI_PATH=$(find ~/Downloads -name "gui_merge.py" 2>/dev/null | head -1)
    if [ -n "$GUI_PATH" ]; then
        cd "$(dirname "$GUI_PATH")"
    else
        echo "No se encontró gui_merge.py"
        echo "Presiona Enter para cerrar..."
        read
        exit 1
    fi
fi

# Ejecutar la GUI
python3 gui_merge.py

# Mantener ventana abierta si hay error
if [ $? -ne 0 ]; then
    echo ""
    echo "Presiona Enter para cerrar..."
    read
fi
EOF

# Hacer ejecutable
chmod +x "$DESKTOP/Merge F42.command"

# Copiar también el archivo Python al escritorio (opcional)
cp "$SCRIPT_DIR/gui_merge.py" "$DESKTOP/" 2>/dev/null

echo "✅ Acceso directo creado en el escritorio: 'Merge F42.command'"
echo ""
echo "Para usarlo:"
echo "  1. Doble clic en 'Merge F42.command' en tu escritorio"
echo "  2. Si sale advertencia de seguridad: clic derecho > Abrir"
echo ""
echo "El archivo F42_MERGED.xlsx se guardará en el escritorio"

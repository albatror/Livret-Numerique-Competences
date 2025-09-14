#!/bin/bash
set -e

PY_VER="3.11.0"
WORKDIR="$(cd "$(dirname "$0")" && pwd)"
PY_DIR="$WORKDIR/python_embed"
PY_TAR="$WORKDIR/python-$PY_VER-macosx.tar.xz"  # On va chercher le binaire tar.xz
PY_URL="https://www.python.org/ftp/python/$PY_VER/python-$PY_VER-macos11.pkg"
INTERFACE_SCRIPT="$WORKDIR/Interface.py"
REQ_FILE="$WORKDIR/requirements.txt"

echo "=== Vérification Python ==="
if command -v python3 >/dev/null 2>&1; then
    echo "Python système détecté"
    PYTHON="python3"
else
    echo "Python non trouvé. Installation portable..."
    # Téléchargement du binaire (pkg ou tar.gz). Sous macOS il n’y a pas d’embed, on installe localement.
    if [ ! -d "$PY_DIR" ]; then
        echo "Téléchargement Python $PY_VER..."
        curl -L -o "$PY_TAR" "https://www.python.org/ftp/python/$PY_VER/Python-$PY_VER.tgz"
        echo "Décompression..."
        mkdir -p "$PY_DIR"
        tar -xzf "$PY_TAR" -C "$PY_DIR" --strip-components=1
    fi
    PYTHON="$PY_DIR/bin/python3"
fi

echo
echo "=== Vérification pip ==="
if ! $PYTHON -m pip --version >/dev/null 2>&1; then
    echo "Installation pip..."
    $PYTHON -m ensurepip --default-pip || {
        curl -L https://bootstrap.pypa.io/get-pip.py -o "$WORKDIR/get-pip.py"
        $PYTHON "$WORKDIR/get-pip.py"
        rm "$WORKDIR/get-pip.py"
    }
fi

echo "Mise à jour pip/setuptools/wheel..."
$PYTHON -m pip install --upgrade pip setuptools wheel

if [ -f "$REQ_FILE" ]; then
    echo "Installation des dépendances requirements.txt..."
    $PYTHON -m pip install -r "$REQ_FILE"
fi

echo
echo "=== Lancement Interface.py ==="
$PYTHON "$INTERFACE_SCRIPT"

echo "Nettoyage fichiers temporaires..."
[ -f "$PY_TAR" ] && rm -f "$PY_TAR"

echo
echo "=== Terminé ==="

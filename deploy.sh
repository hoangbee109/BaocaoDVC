#!/bin/bash
set -e

# === CAU HINH ===
VPS_HOST="catpt.tech"
VPS_PORT="22"
VPS_USER="nmt"
VPS_KEY="/Users/tuong/Documents/id_rsa"  # Doi duong dan private key neu can
REMOTE_DIR="/opt/baocaohcc"
ARCHIVE="baocaohcc.tar.gz"

# === MAU SAC ===
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

log()  { echo -e "${GREEN}[OK]${NC} $1"; }
warn() { echo -e "${YELLOW}[!]${NC} $1"; }
err()  { echo -e "${RED}[ERR]${NC} $1"; exit 1; }

SSH_CMD="ssh -i $VPS_KEY -p $VPS_PORT -o StrictHostKeyChecking=no $VPS_USER@$VPS_HOST"
SCP_CMD="scp -i $VPS_KEY -P $VPS_PORT -o StrictHostKeyChecking=no"

# === KIEM TRA KET NOI ===
echo ""
echo "========================================="
echo "  DEPLOY HE THONG BAO CAO -> $VPS_HOST"
echo "========================================="
echo ""

warn "Kiem tra ket noi SSH..."
$SSH_CMD "echo 'OK'" > /dev/null 2>&1 || err "Khong ket noi duoc toi $VPS_USER@$VPS_HOST:$VPS_PORT. Kiem tra lai private key va thong tin VPS."
log "Ket noi SSH thanh cong."

# === DONG GOI ===
warn "Dong goi project..."
tar czf "$ARCHIVE" \
    --exclude='node_modules' \
    --exclude='.git' \
    --exclude='npm-debug.log' \
    --exclude='.env' \
    --exclude="$ARCHIVE" \
    .
log "Da tao $ARCHIVE ($(du -h "$ARCHIVE" | cut -f1))"

# === UPLOAD ===
warn "Upload len VPS..."
$SSH_CMD "sudo mkdir -p $REMOTE_DIR && sudo chown $VPS_USER:$VPS_USER $REMOTE_DIR"
$SCP_CMD "$ARCHIVE" "$VPS_USER@$VPS_HOST:$REMOTE_DIR/$ARCHIVE"
log "Upload thanh cong."

# === GIAI NEN & DEPLOY ===
warn "Giai nen va khoi dong Docker..."
$SSH_CMD << ENDSSH
    cd $REMOTE_DIR
    tar xzf $ARCHIVE
    rm -f $ARCHIVE

    # Cai docker-compose neu chua co
    if ! command -v docker &> /dev/null; then
        echo "Docker chua cai dat. Vui long cai Docker truoc."
        exit 1
    fi

    # Dung container cu, build lai va chay
    docker compose down || docker-compose down || true
    docker compose up -d --build || docker-compose up -d --build

    echo ""
    echo "=== TRANG THAI CONTAINER ==="
    docker compose ps 2>/dev/null || docker-compose ps 2>/dev/null
ENDSSH

# === DON DEP LOCAL ===
rm -f "$ARCHIVE"
log "Da xoa file tam local."

echo ""
echo "========================================="
echo -e "  ${GREEN}DEPLOY THANH CONG!${NC}"
echo "  URL: http://$VPS_HOST:3000"
echo "========================================="
echo ""

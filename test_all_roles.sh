#!/bin/bash
BASE="https://prototypepenjualan.up.railway.app"

# Login function - returns cookie file
do_login() {
    local user=$1 pass=$2 role=$3
    local cookie="/tmp/cookie_${role}.txt"
    rm -f "$cookie"
    
    curl -s -c "$cookie" -b "$cookie" -X POST "$BASE/login" \
        -d "username=$user&password=$pass" \
        -L -o /dev/null -w "%{http_code}" 2>/dev/null
    echo ""
}

# Test function
test_route() {
    local role=$1 route=$2
    local cookie="/tmp/cookie_${role}.txt"
    local code=$(curl -s -b "$cookie" -o /dev/null -w "%{http_code}" "$BASE$route" 2>/dev/null)
    if [ "$code" = "200" ]; then
        echo "✅ $code $route"
    elif [ "$code" = "307" ] || [ "$code" = "302" ] || [ "$code" = "303" ]; then
        echo "↪️  $code $route (redirect)"
    else
        echo "❌ $code $route"
    fi
}

echo "═══════════════════════════════════"
echo "  LOGIN ALL ROLES"
echo "═══════════════════════════════════"
echo -n "MASTER: "; do_login "master-dev" "masterozi" "master"
echo -n "BOS:    "; do_login "admin" "admin123" "bos"
echo -n "OG:     "; do_login "og1" "og123" "og"
echo -n "KARYAWAN: "; do_login "kasir1" "emp123" "karyawan"

echo ""
echo "═══════════════════════════════════"
echo "  TESTING MASTER ROLE"
echo "═══════════════════════════════════"
ROUTES=(
    "/" "/produk" "/stok" "/stok/masuk" "/stok/keluar"
    "/penjualan" "/pelanggan" "/supplier" "/hutang"
    "/opname" "/opname/riwayat" "/laporan" "/laporan/kasir"
    "/closing" "/closing/riwayat" "/garansi"
    "/finance/kas" "/finance/cashflow"
    "/brand" "/serial" "/retur" "/po"
    "/kategori" "/service" "/promo" "/komisi"
    "/users" "/notifikasi"
    "/master/features" "/master/clients"
    "/pengaturan" "/backup" "/audit-log"
)
for r in "${ROUTES[@]}"; do
    test_route "master" "$r"
done

echo ""
echo "═══════════════════════════════════"
echo "  TESTING BOS ROLE"
echo "═══════════════════════════════════"
for r in "${ROUTES[@]}"; do
    [[ "$r" == /master/* ]] && continue
    test_route "bos" "$r"
done

echo ""
echo "═══════════════════════════════════"
echo "  TESTING OG ROLE"
echo "═══════════════════════════════════"
for r in "${ROUTES[@]}"; do
    [[ "$r" == /master/* ]] && continue
    test_route "og" "$r"
done

echo ""
echo "═══════════════════════════════════"
echo "  TESTING KARYAWAN ROLE"
echo "═══════════════════════════════════"
for r in "${ROUTES[@]}"; do
    [[ "$r" == /master/* ]] && continue
    test_route "karyawan" "$r"
done

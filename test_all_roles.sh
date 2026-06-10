#!/bin/bash
BASE="https://prototypepenjualan.up.railway.app"

do_login() {
    local user=$1 pass=$2 role=$3
    local cookie="/tmp/cookie_${role}.txt"
    rm -f "$cookie"
    # Login - don't follow redirect, just capture cookie
    local code=$(curl -s -c "$cookie" -X POST "$BASE/login" \
        -d "username=$user&password=$pass" \
        -o /dev/null -w "%{http_code}" 2>/dev/null)
    echo "login=$code"
}

test_route() {
    local role=$1 route=$2
    local cookie="/tmp/cookie_${role}.txt"
    local tmpfile="/tmp/resp_${role}_$$.html"
    local code=$(curl -s -b "$cookie" -o "$tmpfile" -w "%{http_code}" "$BASE$route" 2>/dev/null)
    local size=$(wc -c < "$tmpfile" 2>/dev/null || echo 0)
    rm -f "$tmpfile"
    
    if [ "$code" = "200" ]; then
        if [ "$size" -lt 500 ]; then
            echo "⚠️  $code ($size bytes) $route"
        else
            echo "✅ $code ($size bytes) $route"
        fi
    elif [ "$code" = "303" ] || [ "$code" = "302" ] || [ "$code" = "307" ]; then
        echo "↪️  $code $route (redirect=not permitted)"
    elif [ "$code" = "403" ]; then
        echo "🔒 $code $route (forbidden)"
    elif [ "$code" = "422" ]; then
        echo "❌ $code $route (validation error)"
    elif [ "$code" = "500" ]; then
        echo "💥 $code $route (server error!)"
    else
        echo "❌ $code $route"
    fi
}

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

echo "═══════════════════════════════════════════"
echo "  COMPREHENSIVE ROLE TEST"
echo "═══════════════════════════════════════════"
echo ""

for role_info in "master:master-dev:masterozi" "bos:admin:admin123" "og:og1:og123" "karyawan:kasir1:emp123"; do
    IFS=':' read -r role user pass <<< "$role_info"
    echo "═══ $role ($user) ═══"
    do_login "$user" "$pass" "$role"
    
    for r in "${ROUTES[@]}"; do
        # Skip master routes for non-master roles
        if [[ "$role" != "master" && "$r" == /master/* ]]; then
            continue
        fi
        test_route "$role" "$r"
    done
    echo ""
done

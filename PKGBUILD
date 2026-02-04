# Maintainer: sudacode <sudacode@example.com>
pkgname=subminer
pkgver=1.0.0
pkgrel=1
pkgdesc="All-in-one sentence mining overlay with AnkiConnect and dictionary integration"
arch=('x86_64')
url="https://github.com/sudacode/subminer"
license=('GPL-3.0-or-later')
depends=(
    'mpv'
    'mecab'
    'mecab-ipadic'
    'fuse2'
)
optdepends=(
    'fzf: Terminal-based video file picker'
    'rofi: GUI-based video file picker'
    'chafa: Video thumbnail previews in fzf'
    'ffmpegthumbnailer: Generate video thumbnails'
)
source=(
    "$pkgname-$pkgver.AppImage::$url/releases/download/v$pkgver/subminer-$pkgver.AppImage"
    "subminer-$pkgver::$url/releases/download/v$pkgver/subminer"
    "catppuccin-macchiato.rasi::$url/raw/v$pkgver/catppuccin-macchiato.rasi"
)
sha256sums=('SKIP' 'SKIP' 'SKIP')

package() {
    install -Dm755 "$pkgname-$pkgver.AppImage" "$pkgdir/opt/$pkgname/subminer.AppImage"
    install -Dm755 "subminer-$pkgver" "$pkgdir/usr/bin/subminer"
    install -Dm644 catppuccin-macchiato.rasi "$pkgdir/opt/$pkgname/catppuccin-macchiato.rasi"

    install -d "$pkgdir/usr/bin"
    ln -s "/opt/$pkgname/subminer.AppImage" "$pkgdir/usr/bin/subminer"
}

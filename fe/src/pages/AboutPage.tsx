import {
  ArrowRight,
  Code2,
  Database,
  FileSpreadsheet,
  Files,
  HardDrive,
  ShieldCheck,
  Workflow,
} from 'lucide-react'
import type { PageKey } from '../components/Sidebar'

type AboutPageProps = {
  onNavigate: (page: PageKey) => void
}

export default function AboutPage({
  onNavigate,
}: AboutPageProps) {
  return (
    <main className="page-content about-page">
      {/* Header halaman */}
      <section className="page-intro">
        <div>
          <p className="eyebrow">WELDING DEPARTMENT</p>

          <h1>Tentang Aplikasi</h1>

          <p className="page-description">
            FSS Control Center, tujuan pengembangannya,
            dan siapa yang membangunnya.
          </p>
        </div>

        <span className="about-version-badge">
          Versi 0.1.0 · Pengembangan
        </span>
      </section>

      {/* Pengenalan aplikasi */}
      <section className="about-hero">
        <div className="about-hero-content">
          <span className="about-kicker">
            FSS CONTROL CENTER
          </span>

          <h2>
            Pengolahan FSS yang lebih sederhana,
            terstruktur, dan terukur.
          </h2>

          <p>
            FSS Control Center dikembangkan untuk membantu
            mengotomatisasi penggabungan file FSS customer
            dan pembaruan data FSS pada workbook schedule Excel.
            Tujuannya adalah mengurangi pekerjaan copy-paste
            manual dan membantu menjaga ketepatan data.
          </p>

          <div className="about-hero-tags">
            <span>
              <FileSpreadsheet size={15} />
              Microsoft Excel
            </span>

            <span>
              <HardDrive size={15} />
              Aplikasi lokal
            </span>
          </div>
        </div>

        <div className="about-hero-logo">
          <div className="about-logo-square">
            <FileSpreadsheet size={48} />
          </div>

          <span>Welding Department</span>
        </div>
      </section>

      {/* Fitur utama */}
      <section className="about-section">
        <div className="about-section-heading">
          <p className="eyebrow">FITUR UTAMA</p>

          <h2>Dua pekerjaan inti dalam satu aplikasi</h2>

          <p>
            Aplikasi difokuskan pada dua proses utama agar
            alur kerja mudah dipahami dan diperiksa.
          </p>
        </div>

        <div className="about-feature-grid">
          <article className="about-feature-card">
            <div className="about-feature-icon orange">
              <Files size={24} />
            </div>

            <span className="about-feature-number">01</span>

            <h3>Build Data FSS</h3>

            <p>
              Menggabungkan file Firm dan Forecast menjadi
              satu master dengan struktur 13 kolom yang
              telah ditentukan.
            </p>

            <button
              type="button"
              className="about-feature-link"
              onClick={() => onNavigate('build')}
            >
              Buka halaman Build
              <ArrowRight size={16} />
            </button>
          </article>

          <article className="about-feature-card">
            <div className="about-feature-icon purple">
              <Workflow size={24} />
            </div>

            <span className="about-feature-number">02</span>

            <h3>Update Schedule DATA</h3>

            <p>
              Menyiapkan rekonsiliasi Firm lama, Firm baru,
              dan Forecast terbaru sambil menjaga integritas
              workbook schedule.
            </p>

            <button
              type="button"
              className="about-feature-link"
              onClick={() => onNavigate('update')}
            >
              Buka halaman Update
              <ArrowRight size={16} />
            </button>
          </article>
        </div>
      </section>

      {/* Prinsip aplikasi */}
      <section className="about-section">
        <div className="about-section-heading">
          <p className="eyebrow">PRINSIP PENGEMBANGAN</p>

          <h2>Dirancang dengan mengutamakan keandalan data</h2>

          <p>
            Pengolahan data harus mengikuti aturan yang jelas
            dan hasilnya harus bisa diperiksa.
          </p>
        </div>

        <div className="about-principles">
          <article className="about-principle-row">
            <div className="about-principle-icon">
              <ShieldCheck size={21} />
            </div>

            <div>
              <h3>Validasi sebelum hasil dinyatakan berhasil</h3>

              <p>
                Struktur kolom, identitas order, perubahan data,
                dan hasil penulisan perlu diperiksa. Order yang
                ambigu tidak boleh dihapus secara otomatis.
              </p>
            </div>
          </article>

          <article className="about-principle-row">
            <div className="about-principle-icon">
              <HardDrive size={21} />
            </div>

            <div>
              <h3>Pemrosesan lokal</h3>

              <p>
                Arsitektur dirancang agar file customer diproses
                pada komputer pengguna, bukan diunggah ke server
                cloud eksternal. Integrasi backend lokal masih
                dalam tahap pengembangan.
              </p>
            </div>
          </article>

          <article className="about-principle-row">
            <div className="about-principle-icon">
              <Database size={21} />
            </div>

            <div>
              <h3>Excel tetap menjadi format kerja utama</h3>

              <p>
                Aplikasi membantu mengotomatisasi pekerjaan
                pengolahan data tanpa bermaksud menggantikan
                formula production planning di Excel.
              </p>
            </div>
          </article>
        </div>
      </section>

      {/* Profil developer */}
      <section className="about-developer-panel">
        <div className="about-developer-heading">
          <p className="eyebrow">DI BALIK APLIKASI</p>

          <h2>Tentang Developer</h2>

          <p>
            Aplikasi ini dikembangkan untuk mempermudah
            pekerjaan pengolahan FSS dan pembaruan schedule
            di lingkungan Welding Department.
          </p>
        </div>

        <div className="about-developer-card">
          <div className="about-developer-avatar">
            <Code2 size={45} />
          </div>
          <div className="about-developer-info">
            <span className="about-developer-role">
              APPLICATION DEVELOPER
            </span>

            <h3>Ahmad Galih Saputro</h3>

            <p>
              Pengembang FSS Control Center. Berfokus pada
              pembuatan alat bantu kerja yang praktis,
              pengolahan data yang terstruktur, dan antarmuka
              aplikasi yang mudah digunakan.
            </p>

            <div className="about-developer-skills">
              <span>
                <Code2 size={14} />
                Application Development
              </span>

              <span>
                <FileSpreadsheet size={14} />
                Excel Workflow Automation
              </span>
            </div>
          </div>
        </div>
      </section>

      {/* Informasi versi */}
      <section className="about-footer-card">
        <div>
          <span className="about-footer-label">
            INFORMASI APLIKASI
          </span>

          <strong>FSS Control Center</strong>

          <p>Welding Department · Versi 0.1.0</p>
        </div>

        <span className="about-development-status">
          Dalam pengembangan
        </span>
      </section>

      <p className="about-copyright">
        FSS Control Center · Welding Department
      </p>
    </main>
  )
}
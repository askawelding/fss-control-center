import { useState } from 'react'

import {
  Check,
  Database,
  FolderCog,
  Info,
  Palette,
  RotateCcw,
  Save,
  Server,
  Settings2,
  ShieldCheck,
} from 'lucide-react'
import wIconWhite from '../assets/w_icon_white.svg'

type TargetSheet = 'DATE' | 'DATA'

export default function SettingsPage() {
  const [targetSheet, setTargetSheet] =
    useState<TargetSheet>('DATE')

  const [showSaveMessage, setShowSaveMessage] =
    useState(false)

  function handleApplySettings() {
    setShowSaveMessage(true)
  }

  function handleReset() {
    setTargetSheet('DATE')
    setShowSaveMessage(false)
  }

  return (
    <main className="page-content settings-page">
      <section className="page-intro">
        <div>
          <p className="eyebrow">WELDING DEPARTMENT</p>

          <h1>Pengaturan</h1>

          <p className="page-description">
            Kelola preferensi proses dan informasi aplikasi lokal.
          </p>
        </div>
      </section>

      <section className="settings-notice">
        <Info size={19} />

        <p>
          Pengaturan pada tahap ini masih berupa demonstrasi
          frontend. Belum ada konfigurasi yang disimpan secara
          permanen ke database atau sistem lokal.
        </p>
      </section>

      {/* Bagian 1: Preferensi proses */}
      <section className="settings-panel">
        <div className="settings-panel-heading">
          <div className="settings-icon orange">
            <Settings2 size={21} />
          </div>

          <div>
            <h2>Preferensi proses</h2>

            <p>
              Tentukan nilai awal yang akan digunakan pada halaman
              Update Schedule.
            </p>
          </div>
        </div>

        <div className="settings-field">
          <div className="settings-field-copy">
            <label htmlFor="default-sheet">
              Worksheet default
            </label>

            <p>
              Worksheet yang dipilih secara awal saat halaman
              Update Schedule dibuka.
            </p>
          </div>

          <select
            id="default-sheet"
            className="settings-select"
            value={targetSheet}
            onChange={(event) => {
              setTargetSheet(
                event.currentTarget.value as TargetSheet,
              )

              setShowSaveMessage(false)
            }}
          >
            <option value="DATE">DATE</option>
            <option value="DATA">DATA</option>
          </select>
        </div>

        <div className="settings-field">
          <div className="settings-field-copy">
            <label>Mode pembaruan</label>

            <p>
              Mode lengkap direkomendasikan untuk memastikan kalkulasi
              dan validasi yang dibutuhkan telah selesai.
            </p>
          </div>

          <div className="settings-readonly-value">
            <ShieldCheck size={17} />
            Update Lengkap
            <span>Default</span>
          </div>
        </div>

        <div className="settings-help">
          <Info size={16} />

          <p>
            Update Cepat belum diaktifkan. Mode tersebut baru boleh
            digunakan jika implementasi dan kebutuhan kalkulasinya
            sudah diperiksa dengan benar.
          </p>
        </div>

        <div className="settings-actions">
          <button
            type="button"
            className="settings-reset-button"
            onClick={handleReset}
          >
            <RotateCcw size={16} />
            Kembalikan default
          </button>

          <button
            type="button"
            className="primary-button"
            onClick={handleApplySettings}
          >
            <Save size={16} />
            Terapkan pilihan
          </button>
        </div>

        {showSaveMessage && (
          <div className="settings-feedback" role="status">
            <Check size={17} />

            <span>
              Pilihan worksheet {targetSheet} diterapkan pada tampilan
              pengaturan ini. Belum tersimpan secara permanen.
            </span>
          </div>
        )}
      </section>

      {/* Bagian 2: Lokasi dan output */}
      <section className="settings-panel">
        <div className="settings-panel-heading">
          <div className="settings-icon purple">
            <FolderCog size={21} />
          </div>

          <div>
            <h2>Lokasi dan output</h2>

            <p>
              Informasi mengenai penyimpanan file hasil pengolahan.
            </p>
          </div>
        </div>

        <div className="settings-info-row">
          <div>
            <strong>Format nama master</strong>

            <p>
              Nama berdasarkan tanggal rilis, misalnya 20261028.xlsx.
            </p>
          </div>

          <code>YYYYMMDD.xlsx</code>
        </div>

        <div className="settings-info-row">
          <div>
            <strong>Folder output</strong>

            <p>
              Lokasi penyimpanan akan dikelola melalui mekanisme
              backend lokal.
            </p>
          </div>

          <span className="settings-pending-badge">
            Belum dikonfigurasi
          </span>
        </div>

        <div className="settings-info-row">
          <div>
            <strong>File sumber</strong>

            <p>
              File asli harus tetap utuh selama proses pengolahan.
            </p>
          </div>

          <span className="settings-safe-badge">
            <ShieldCheck size={14} />
            Aturan keamanan
          </span>
        </div>
      </section>

      {/* Bagian 3: Status sistem */}
      <section className="settings-panel">
        <div className="settings-panel-heading">
          <div className="settings-icon gray">
            <Server size={21} />
          </div>

          <div>
            <h2>Status sistem</h2>

            <p>
              Status koneksi akan diperoleh dari backend setelah
              integrasi dilakukan.
            </p>
          </div>
        </div>

        <div className="settings-system-grid">
          <div className="settings-system-card">
            <div className="settings-system-top">
              <Server size={19} />

              <span className="settings-system-dot" />
            </div>

            <strong>Backend lokal</strong>

            <span>Belum terhubung</span>
          </div>

          <div className="settings-system-card">
            <div className="settings-system-top">
              <Settings2 size={19} />

              <span className="settings-system-dot" />
            </div>

            <strong>Microsoft Excel</strong>

            <span>Belum diperiksa</span>
          </div>

          <div className="settings-system-card">
            <div className="settings-system-top">
              <Database size={19} />

              <span className="settings-system-dot" />
            </div>

            <strong>Database riwayat</strong>

            <span>Belum tersedia</span>
          </div>
        </div>
      </section>

      {/* Bagian 4: Backup */}
      <section className="settings-panel">
        <div className="settings-panel-heading">
          <div className="settings-icon gray">
            <Database size={21} />
          </div>

          <div>
            <h2>Data dan backup</h2>

            <p>
              Perlindungan database metadata riwayat proses.
            </p>
          </div>
        </div>

        <div className="settings-backup-row">
          <div>
            <strong>Backup database riwayat</strong>

            <p>
              Fitur ini akan tersedia setelah SQLite dan layanan
              backup lokal selesai dibuat.
            </p>
          </div>

          <button
            type="button"
            className="settings-disabled-button"
            disabled
          >
            <Database size={15} />
            Belum tersedia
          </button>
        </div>
      </section>

      {/* Bagian 5: Branding */}
      <section className="settings-panel">
        <div className="settings-panel-heading">
          <div className="settings-icon orange">
            <Palette size={21} />
          </div>

          <div>
            <h2>Tampilan dan branding</h2>

            <p>
              Identitas visual yang digunakan pada frontend.
            </p>
          </div>
        </div>

        <div className="settings-brand-row">
          <div className="settings-brand-logo">
            <img src={wIconWhite} alt="FSS Control Center" 
            style={{ width: '75%', height: '75%' }} />
          </div>

          <div className="settings-brand-copy">
            <strong>FSS Control Center</strong>

            <p>Welding Department</p>

            <span>Font: Poppins</span>
          </div>
        </div>

        <div className="settings-color-row">
          <div
            className="settings-color-swatch"
            aria-hidden="true"
          />

          <div>
            <strong>Primary Orange</strong>

            <p>Warna utama aplikasi</p>
          </div>

          <code>#F47721</code>
        </div>

        <div className="settings-info-row">
          <div>
            <strong>Versi frontend</strong>

            <p>Versi pengembangan antarmuka saat ini.</p>
          </div>

          <code>0.1.0</code>
        </div>
      </section>
    </main>
  )
}
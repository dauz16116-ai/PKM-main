from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from app.models.user import User
from app.models.siswa import Siswa
from app import db
from app.utils.decorators import role_required
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
import io, os, re

admin_bp = Blueprint('admin', __name__, url_prefix='/admin')

def extract_nama_depan(nama):
    if not nama:
        return 'guru'
    titles = {'ustadz', 'ustadzah', 'ust', 'ustzh', 'guru', 'bapak', 'ibu', 'pak', 'bu', 'dr', 'drs', 'dra', 'h', 'hj', 'ir', 'mr', 'mrs', 'ms'}
    words = [re.sub(r'[^a-zA-Z0-9]', '', w) for w in str(nama).split()]
    clean_words = [w.lower() for w in words if w and w.lower() not in titles]
    if clean_words:
        return clean_words[0]
    if words and words[0]:
        return words[0].lower()
    return 'guru'

def parse_kelas_list(kelas_str):
    if not kelas_str or kelas_str.strip() == '-':
        return []
    parts = kelas_str.replace(';', ',').replace('/', ',').split(',')
    cleaned = []
    for p in parts:
        c = p.strip().upper()
        if c and c != '-' and c not in cleaned:
            cleaned.append(c)
    return cleaned

def get_assigned_classes_map(exclude_guru_id=None):
    assigned = {}
    query = User.query.filter_by(role='guru', status='Aktif')
    if exclude_guru_id:
        query = query.filter(User.id != exclude_guru_id)
    
    for g in query.all():
        classes = parse_kelas_list(g.kelas)
        for c in classes:
            assigned[c] = g.nama or g.username
    return assigned

@admin_bp.route('/dashboard')
@role_required('admin')
def dashboard():
    guru_list = User.query.filter_by(role='guru').all()
    siswa_list = Siswa.query.all()
    return render_template('admin/dashboard_admin.html', guru_list=guru_list, siswa_list=siswa_list)

@admin_bp.route('/template_guru')
@role_required('admin')
def template_guru():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template_Guru"

    img_path = os.path.join('app', 'static', 'img', 'logo-sdit-sm.png')
    if not os.path.exists(img_path):
        img_path = os.path.join('PKM-main', 'app', 'static', 'img', 'logo-sdit-sm.png')
    if os.path.exists(img_path):
        try:
            img = ExcelImage(img_path)
            img.width = 52
            img.height = 52
            ws.add_image(img, 'A1')
        except Exception:
            pass

    ws.merge_cells('B1:D1')
    ws['B1'] = "SDIT BINA MULIA"
    ws['B1'].font = Font(name="Arial", size=14, bold=True, color="0F172A")

    ws.merge_cells('B2:D2')
    ws['B2'] = "FORMAT RESMI IMPORT DATA GURU & MUSYRIF"
    ws['B2'].font = Font(name="Arial", size=10, bold=True, color="334155")

    ws['B3'] = "Tahun Ajaran: 2026/2027"
    ws['B3'].font = Font(name="Arial", size=9, bold=True, color="475569")
    ws['D3'] = "Kelas Binaan: [Contoh: 1A, 2B]"
    ws['D3'].font = Font(name="Arial", size=9, bold=True, color="475569")

    ws.merge_cells('B4:D4')
    ws['B4'] = "Catatan: ID Guru dijadikan password login awal. Username otomatis dibuat dari nama depan guru."
    ws['B4'].font = Font(name="Arial", size=8, italic=True, color="64748B")

    ws.append([]) # row 5 blank

    header_fill = PatternFill(start_color="0f172a", end_color="0f172a", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    headers = ["No", "ID Guru", "Nama Lengkap", "Kelas Binaan"]
    ws.append(headers) # row 6

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=6, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    ws.row_dimensions[6].height = 26

    # Kosongan: Buat baris kosong siap diinput tanpa data dummy
    for row_idx in range(7, 22):
        ws.row_dimensions[row_idx].height = 22
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx == 1:
                cell.value = row_idx - 6
                cell.alignment = center_align
            elif col_idx in [2, 4]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    widths = [6, 18, 36, 24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="Template_Resmi_Import_Guru_SDIT_Bina_Mulia.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@admin_bp.route('/tambah_guru', methods=['POST'])
@role_required('admin')
def tambah_guru():
    nama = request.form.get('nama', '').strip()
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()
    kelas_input = request.form.get('kelas', '-').strip()
    status = request.form.get('status', 'Aktif')

    if not nama or not username or not password:
        flash('Semua kolom wajib diisi!', 'danger')
        return redirect(url_for('admin.dashboard'))

    if User.query.filter(db.func.lower(User.username) == username.lower()).first():
        flash(f'Username "{username}" sudah digunakan oleh akun lain!', 'danger')
        return redirect(url_for('admin.dashboard'))

    kelas_list = parse_kelas_list(kelas_input)
    if status == 'Aktif' and kelas_list:
        assigned_map = get_assigned_classes_map()
        conflict = []
        for c in kelas_list:
            if c in assigned_map:
                conflict.append(f"Kelas {c} (sudah dibina oleh {assigned_map[c]})")
        
        if conflict:
            flash(f'Gagal menambahkan guru! {", ".join(conflict)}. Setiap kelas binaan hanya boleh dibina oleh 1 guru!', 'danger')
            return redirect(url_for('admin.dashboard'))

    formatted_kelas = ", ".join(kelas_list) if kelas_list else '-'
    new_guru = User(nama=nama, username=username, role='guru', kelas=formatted_kelas, status=status)
    new_guru.set_password(password)
    db.session.add(new_guru)
    db.session.commit()
    flash(f'Guru {nama} dengan kelas {formatted_kelas} berhasil ditambahkan!', 'success')

    return redirect(url_for('admin.dashboard'))

@admin_bp.route('/edit_guru/<int:guru_id>', methods=['POST'])
@role_required('admin')
def edit_guru(guru_id):
    guru = User.query.get_or_404(guru_id)
    nama = request.form.get('nama', '').strip()
    username = request.form.get('username', '').strip()
    kelas_input = request.form.get('kelas', '-').strip()
    status = request.form.get('status', 'Aktif')

    existing_user = User.query.filter(
        (db.func.lower(User.username) == username.lower()) & 
        (User.id != guru_id)
    ).first()
    if existing_user:
        flash(f'Username "{username}" sudah digunakan guru lain!', 'danger')
        return redirect(url_for('admin.dashboard'))

    kelas_list = parse_kelas_list(kelas_input)
    if status == 'Aktif' and kelas_list:
        assigned_map = get_assigned_classes_map(exclude_guru_id=guru.id)
        conflict = []
        for c in kelas_list:
            if c in assigned_map:
                conflict.append(f"Kelas {c} (sudah dibina oleh {assigned_map[c]})")
        
        if conflict:
            flash(f'Gagal mengubah data guru! {", ".join(conflict)}. Setiap kelas binaan hanya boleh dibina oleh 1 guru!', 'danger')
            return redirect(url_for('admin.dashboard'))

    guru.nama = nama
    guru.username = username
    guru.kelas = ", ".join(kelas_list) if kelas_list else '-'
    guru.status = status

    password_baru = request.form.get('password', '').strip()
    if password_baru:
        guru.set_password(password_baru)

    db.session.commit()
    flash(f'Data guru {guru.nama} berhasil diperbarui!', 'success')
    return redirect(url_for('admin.dashboard'))

@admin_bp.route('/hapus_guru/<int:guru_id>')
@role_required('admin')
def hapus_guru(guru_id):
    guru = User.query.get_or_404(guru_id)
    nama = guru.nama
    db.session.delete(guru)
    db.session.commit()
    flash(f'Akun guru {nama} berhasil dihapus dari sistem!', 'success')
    return redirect(url_for('admin.dashboard'))

@admin_bp.route('/edit_siswa/<id_siswa>', methods=['POST'])
@role_required('admin')
def edit_siswa(id_siswa):
    siswa = Siswa.query.get_or_404(id_siswa)
    nama = request.form.get('nama', '').strip()
    username = request.form.get('username', '').strip()
    kelas = request.form.get('kelas', '').strip().upper()
    status = request.form.get('status', 'Aktif')

    if not nama or not username:
        flash('Nama dan Username murid tidak boleh kosong!', 'danger')
        return redirect(url_for('admin.dashboard'))

    ortu = None
    if siswa.user_id_ortu:
        ortu = User.query.get(siswa.user_id_ortu)
    if not ortu:
        ortu = User.query.filter_by(username=siswa.nama, role='ortu').first()

    # Cek username conflict
    conflict_user = User.query.filter(
        (db.func.lower(User.username) == username.lower()) &
        (User.id != (ortu.id if ortu else 0))
    ).first()
    if conflict_user:
        flash(f'Username "{username}" sudah digunakan pengguna lain!', 'danger')
        return redirect(url_for('admin.dashboard'))

    siswa.nama = nama
    siswa.kelas = kelas
    siswa.status = status

    if not ortu:
        ortu = User(username=username, role='ortu', nama=f"Wali {nama}", kelas=kelas, status=status)
        ortu.set_password('1234')
        db.session.add(ortu)
        db.session.flush()
        siswa.user_id_ortu = ortu.id
    else:
        ortu.username = username
        ortu.nama = f"Wali {nama}"
        ortu.kelas = kelas
        ortu.status = status

    db.session.commit()
    flash(f'Data murid {nama} (Username: {username}) berhasil diperbarui!', 'success')
    return redirect(url_for('admin.dashboard'))

@admin_bp.route('/edit_password_siswa/<id_siswa>', methods=['POST'])
@role_required('admin')
def edit_password_siswa(id_siswa):
    siswa = Siswa.query.get_or_404(id_siswa)
    password_baru = request.form.get('password', '').strip()

    if not password_baru:
        flash('Password baru tidak boleh kosong!', 'danger')
        return redirect(url_for('admin.dashboard'))

    ortu = None
    if siswa.user_id_ortu:
        ortu = User.query.get(siswa.user_id_ortu)
    if not ortu:
        ortu = User.query.filter_by(username=siswa.nama, role='ortu').first()

    if not ortu:
        ortu = User(username=siswa.nama, role='ortu', nama=f"Wali {siswa.nama}", kelas=siswa.kelas, status='Aktif')
        ortu.set_password(password_baru)
        db.session.add(ortu)
        db.session.flush()
        siswa.user_id_ortu = ortu.id
    else:
        ortu.set_password(password_baru)

    db.session.commit()
    flash(f'Password akun murid/wali untuk {siswa.nama} berhasil diperbarui!', 'success')
    return redirect(url_for('admin.dashboard'))

@admin_bp.route('/import_guru', methods=['POST'])
@role_required('admin')
def import_guru():
    file = request.files.get('file_excel_guru')
    if file and file.filename.endswith(('.xlsx', '.xls')):
        try:
            # Deteksi baris header secara fleksibel
            df_raw = pd.read_excel(file, header=None)
            header_idx = 0
            for idx, r in df_raw.iterrows():
                row_str = " ".join([str(v).lower() for v in r.values if pd.notna(v)])
                if any(kw in row_str for kw in ['nama', 'guru', 'nip', 'id guru', 'kelas']):
                    header_idx = idx
                    break
            
            file.seek(0)
            df = pd.read_excel(file, skiprows=header_idx)
            imported_count = 0
            skipped = []
            
            assigned_map = get_assigned_classes_map()

            col_map = {str(c).strip().lower(): c for c in df.columns}
            def find_val(row, candidates, fallback=''):
                for cand in candidates:
                    for clean_k, real_col in col_map.items():
                        if cand == clean_k or cand in clean_k:
                            val = str(row.get(real_col, '')).strip()
                            if val and val.lower() != 'nan':
                                return val
                return fallback

            for _, row in df.iterrows():
                nama = find_val(row, ['nama lengkap', 'nama guru', 'nama', 'teacher'])
                if not nama or nama.lower() == 'nan':
                    continue

                raw_id_guru = find_val(row, ['id guru', 'id', 'nip', 'kode guru', 'id_guru', 'password'])
                if raw_id_guru and raw_id_guru.lower() != 'nan':
                    id_guru = re.sub(r'\.0$', '', raw_id_guru)
                else:
                    g_count = User.query.filter_by(role='guru').count() + imported_count + 1
                    id_guru = f"G{g_count:03d}"

                raw_kelas = find_val(row, ['kelas binaan', 'kelas', 'rombel'], '-')

                # Logika: username = nama depan, pass = id guru
                nama_depan = extract_nama_depan(nama)
                username_cand = nama_depan
                if User.query.filter(db.func.lower(User.username) == username_cand.lower()).first():
                    clean_id = re.sub(r'[^a-zA-Z0-9]', '', id_guru).lower()
                    username_cand = f"{nama_depan}_{clean_id}" if clean_id else f"{nama_depan}_{User.query.count() + 1}"

                password = id_guru

                kelas_list = parse_kelas_list(raw_kelas)
                has_conflict = False
                for c in kelas_list:
                    if c in assigned_map:
                        skipped.append(f"{nama} (Kelas {c} sudah dibina {assigned_map[c]})")
                        has_conflict = True
                        break

                if has_conflict:
                    continue

                formatted_kelas = ", ".join(kelas_list) if kelas_list else '-'
                new_guru = User(nama=nama, username=username_cand, role='guru', kelas=formatted_kelas, status='Aktif')
                new_guru.set_password(password)
                db.session.add(new_guru)
                
                for c in kelas_list:
                    assigned_map[c] = nama
                imported_count += 1

            db.session.commit()
            msg = f'Berhasil mengimpor {imported_count} akun guru! (Username: nama depan, Password: ID Guru)'
            if skipped:
                msg += f' Dilewati: {", ".join(skipped)}'
            flash(msg, 'success' if imported_count > 0 else 'warning')
        except Exception as e:
            flash(f'Gagal mengimpor file Excel: {str(e)}', 'danger')
    else:
        flash('Format file tidak valid! Harap unggah file .xlsx atau .xls', 'danger')
    return redirect(url_for('admin.dashboard'))

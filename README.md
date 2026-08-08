# 🇹🇭 Thai ID + Phone Decoder
<p align="center">

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python\&logoColor=white)
![Version](https://img.shields.io/badge/Version-v4.0.0-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)
![Dependencies](https://img.shields.io/badge/Dependencies-Standard%20Library-success)
![License](https://shields.io/badge/license-Apache%202-blue)

</p>

โปรแกรม Python สำหรับวิเคราะห์โครงสร้าง **เลขบัตรประชาชนไทย** และ **เบอร์โทรศัพท์ไทย** โดยใช้ Input เดียว และสามารถตรวจสอบได้อัตโนมัติว่าเป็นเลขประเภทใด

---

## ✨ Features

### 🪪 เลขบัตรประชาชน (Thai ID)

* ตรวจรูปแบบเลข 13 หลัก
* ตรวจสอบ Checksum
* Decode โครงสร้างของเลข
* อ่าน RCODE เพื่อดูจังหวัด / สำนักทะเบียน / พื้นที่ที่เกี่ยวข้อง
* อัปเดตฐานข้อมูล RCODE จาก BORA/DOPA อัตโนมัติ
* รองรับเลขแบบมีขีดและเลขไทย
* ไม่ใช้หรือ Hack ฐานข้อมูลประชาชน

> [!NOTE]
> โปรแกรมวิเคราะห์เฉพาะข้อมูลที่อยู่ในโครงสร้างของตัวเลขและข้อมูลอ้างอิงสาธารณะ ไม่สามารถใช้หาเจ้าของเลขได้

### 📱 เบอร์โทรศัพท์ (Thai Phone)

* รองรับ Mobile / Landline / VoIP / Short Number
* ตรวจ Format และ Normalize อัตโนมัติ
* รองรับรูปแบบ เช่น `0812345678`
* รองรับ `081-234-5678`
* รองรับ International Format เช่น `+66812345678`
* Decode Prefix
* Decode Area Code ของโทรศัพท์บ้าน
* แสดงข้อมูลช่วงเลขและผู้ให้บริการที่เคยได้รับการจัดสรร
* รองรับกรณี Mobile Number Portability (MNP)

> [!WARNING]
> Prefix ไม่สามารถใช้ยืนยันค่ายมือถือปัจจุบันได้ 100% เพราะผู้ใช้งานสามารถย้ายค่ายเบอร์เดิมได้

---

## 🛠️ ข้อมูลโปรเจกต์

| รายการ          | ข้อมูล                  |
| --------------- | ----------------------- |
| Language        | Python                  |
| Python Version  | 3.10+                   |
| Dependencies    | Python Standard Library |
| Platform        | Windows / Linux / macOS |
| Current Version | v4.0.0                  |

ไม่ต้องติดตั้ง Library เพิ่ม

---

## 🚀 วิธีใช้งาน

รูปแบบคำสั่งหลัก:

```bash
python thai_id_phone_decoder.py <NUMBER>
```

โปรแกรมจะตรวจให้อัตโนมัติว่า Input เป็น **Thai ID** หรือ **Thai Phone**

### เบอร์มือถือ

```bash
python thai_id_phone_decoder.py 0812345678
```

หรือ

```bash
python thai_id_phone_decoder.py +66812345678
```

### โทรศัพท์บ้าน

```bash
python thai_id_phone_decoder.py 044213456
```

### เลขบัตรประชาชน

```bash
python thai_id_phone_decoder.py 1409800128861
```

---

## 📦 JSON Output

สำหรับนำ Output ไปใช้กับโปรแกรมอื่นหรือ API:

```bash
python thai_id_phone_decoder.py 0812345678 --json
```

---

## 🔍 Verbose Mode

แสดงรายละเอียดเพิ่มเติมเกี่ยวกับการ Decode:

```bash
python thai_id_phone_decoder.py 1409800128861 --verbose
```

---

## 🥷 Mask

ซ่อนตัวเลขบางส่วน เหมาะสำหรับ Screenshot หรือ Log:

```bash
python thai_id_phone_decoder.py 0812345678 --mask
```

---

## 🎯 กำหนด Input Type เอง

ปกติโปรแกรมจะใช้ `auto`

แต่สามารถบังคับประเภทได้:

```bash
python thai_id_phone_decoder.py 0812345678 --input-type phone
```

หรือ

```bash
python thai_id_phone_decoder.py 1409800128861 --input-type id
```

ประเภทที่รองรับ:

```text
auto
id
phone
```

---

## 🗺️ RCODE

ค้นหา RCODE โดยตรง:

```bash
python thai_id_phone_decoder.py --lookup 3097
```

ค้นจากชื่อพื้นที่:

```bash
python thai_id_phone_decoder.py --find "บัวใหญ่"
```

อัปเดตฐาน RCODE:

```bash
python thai_id_phone_decoder.py --update-only
```

ดูแหล่งข้อมูลที่โปรแกรมใช้อ้างอิง:

```bash
python thai_id_phone_decoder.py --source
```

---

## 🧪 Self Test

ทดสอบระบบภายในโปรแกรม:

```bash
python thai_id_phone_decoder.py --self-test
```

---

## ⚠️ ข้อจำกัด

### เลขบัตรประชาชน

โปรแกรมไม่สามารถใช้เลขบัตรเพื่อค้นหา:

* ชื่อเจ้าของ
* วันเกิด
* อายุ
* ที่อยู่ปัจจุบัน
* ข้อมูลทะเบียนบุคคล
* ข้อมูลจากฐานข้อมูลภาครัฐ

การผ่าน Checksum หมายถึง **โครงสร้างตัวเลขถูกต้อง** ไม่ได้หมายความว่าเลขนั้นมีบุคคลใช้งานจริง

### เบอร์โทรศัพท์

โปรแกรมไม่สามารถยืนยัน:

* เจ้าของเบอร์
* ชื่อบุคคล
* ค่ายมือถือปัจจุบัน 100%
* วันที่เปิด SIM จริง
* ปีที่บุคคลเริ่มใช้เบอร์
* ตำแหน่งปัจจุบันของผู้ใช้

ข้อมูล Operator จาก Prefix เป็นข้อมูลเกี่ยวกับ **ช่วงเลขที่เคยถูกจัดสรร** ไม่ใช่การ Lookup เจ้าของหรือเครือข่ายปัจจุบันแบบ Real-time

---

## 📄 License

โปรเจกต์นี้เผยแพร่ภายใต้ **Apache License**

---

<p align="center">
  <b>Thai ID + Phone Decoder</b><br>
  วิเคราะห์สิ่งที่ตัวเลขสามารถบอกได้ โดยไม่ต้องเข้าถึงฐานข้อมูลส่วนบุคคล
</p>

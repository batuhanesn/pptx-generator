# pptx-generator

Slide verisi + yerel PNG görseller kullanarak PPTX sunum dosyası üreten HTTP servisi. n8n akışındaki "Execute Command (cmd)" node'unun yerini alır: tek bir `POST /generate-pptx` isteğiyle tüm konuların slide verilerini alır ve PPTX üretir.

## Mimari

- **Python + FastAPI** servisi (`main.py`)
- Girdi görselleri **doğrudan diskten** okunur (`FILES_BASE`)
- Üretilen `.pptx` dosyası `OUTPUT_BASE` altına kaydedilir
- Klasör yapısı: `{proje}/Konu{n}/imageSlide{id}.png`
- Slide **metinleri** istek gövdesinde JSON olarak gelir

## Ön koşullar

- Python 3.10+

## Kurulum

```bash
pip install -r requirements.txt
copy .env.example .env   # yolları düzenleyin
python main.py
```

## Yapılandırma (`.env`)

| Değişken | Açıklama |
|---|---|
| `FILES_BASE` | Görsel dosyalarının bulunduğu disk kökü |
| `OUTPUT_BASE` | PPTX çıktısının kaydedileceği klasör |
| `PORT` | HTTP portu (varsayılan 9000) |

## API

### `GET /health`

Servis sağlık kontrolü.

### `POST /generate-pptx`

İstek gövdesi:

```json
{
  "proje_adi": "efeefeefe",
  "konular": [
    {
      "konu_id": 1,
      "slides": [
        { "id": "Slide1", "type": "headline", "content": "Başlık metni" },
        { "id": "Slide2", "type": "bullet",   "content": "Madde Başlığı:\n- Item 1\n- Item 2" }
      ]
    }
  ]
}
```

- Slide tipi `headline`: tek büyük başlık
- Slide tipi `bullet`: ilk satır başlık, sonrakiler madde listesi
- Slide için görsel varsa (`FILES_BASE/{proje}/Konu{n}/image{slideId}.png`) sağ tarafa eklenir

Başarılı yanıt:

```json
{
  "status": "ok",
  "path": "C:/output/efeefeefe/efeefeefe.pptx",
  "slide_count": 4
}
```

### Örnek çağrı

```bash
curl -X POST http://localhost:9000/generate-pptx \
  -H "Content-Type: application/json" \
  -d '{"proje_adi":"test","konular":[{"konu_id":1,"slides":[{"id":"Slide1","type":"headline","content":"Test Başlık"}]}]}'
```

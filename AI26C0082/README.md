# AI26C0082 – Hỏi đáp trên bảng tính Excel

Đọc các file `.xlsx` và `queries.json`, trả lời từng câu hỏi bằng Qwen3.7 Flash qua OpenRouter,
ghi kết quả ra CSV UTF-8 gồm 2 cột `id,answer`.

## Cấu trúc

```
AI26C0082/
├── Dockerfile          ENTRYPOINT ["python", "main.py"]
├── requirements.txt
├── main.py             đọc tham số, chạy song song các câu hỏi, ghi submission.csv
└── src/
    ├── config.py       cấu hình (biến môi trường; tham số dòng lệnh ghi đè)
    ├── parser.py       Excel -> bảng sạch (ô gộp, header nhiều tầng, nhiều bảng/sheet, footnote)
    ├── render.py       bảng -> văn bản cho prompt (giới hạn 20K token / lượt gọi)
    ├── prompts.py      prompt cho đường code, đường đọc bảng, giám khảo
    ├── llm.py          client OpenRouter (bất đồng bộ, retry, đếm token)
    ├── sandbox.py      chạy code pandas do LLM sinh trong tiến trình con, có timeout
    ├── sandbox_runner.py
    ├── solver.py       điều phối 1 câu: 3 lời giải code + 2 lời giải đọc bảng -> bỏ phiếu -> giám khảo
    └── formatter.py    chuẩn hoá đáp án theo định dạng câu hỏi yêu cầu
```

## Build và chạy

```bash
docker build -t ai26c0082 .
docker run --rm \
  -v /path/to/input:/input:ro -v /path/to/output:/output \
  ai26c0082 \
  --xlsx-dir /input/xlsx_files \
  --queries /input/queries.json \
  --output /output/submission.csv \
  --openrouter-key <key> \
  --openrouter-model qwen/qwen3.7-flash \
  --openrouter-base-url https://openrouter.ai/api/v1
```

- API key chỉ nhận qua `--openrouter-key` (hoặc biến môi trường `OPENROUTER_API_KEY`), không có trong mã nguồn.
- Không ghi gì vào `/input`; file tạm của sandbox nằm trong thư mục tạm của hệ thống.
- Nhật ký in ra stdout/stderr. `submission.csv` được ghi lại sau mỗi câu nên luôn hợp lệ, đủ mọi `id`.
- Hạn chót cứng 36 phút (`HARD_DEADLINE`): câu nào chưa xong thì ghi đáp án mặc định rồi thoát mã 0.
- Thoát mã khác 0 nếu thiếu API key hoặc không đọc được `queries.json`.

import cv2
import numpy as np
import matplotlib.pyplot as plt

#box yerleri tekrar ayarlanacak. şuanda hatalı kontur taraması yapmaktadır. kısa sürede düzeltilecek.

STUDENT_BOX_X = 50
STUDENT_BOX_Y = 100
STUDENT_BOX_W = 397   
STUDENT_BOX_H = 515   
NAME_BOX_X = 50
NAME_BOX_Y = 650
NAME_BOX_W = 907
NAME_BOX_H = 931
ANSWER_BOX_X = 500
ANSWER_BOX_Y = 100
ANSWER_BOX_W = 367
ANSWER_BOX_H = 1435
THRESH_VAL = 130
FILL_RATIO_THRESHOLD = 0.3
STUDENT_ROW_COUNT = 10
STUDENT_COL_COUNT = 11
STUDENT_BUBBLE_W = 35
STUDENT_BUBBLE_H = 42
NAME_ROW_COUNT = 30
NAME_COL_COUNT = 30
NAME_BUBBLE_W = 28
NAME_BUBBLE_H = 27
NAME_ROW_H = 31  
ANSWER_ROW_COUNT = 40
ANSWER_COL_COUNT = 6
ANSWER_BUBBLE_W = 58
ANSWER_BUBBLE_H = 32
ANSWER_ROW_H = 35
ANSWER_KEY = "ABBACCCDDBBBBCCCCAAAADDDDABBACCCCABCDABCD"[:40]
def read_omr_form(image_path):
    img = cv2.imread(image_path)
    if img is None:
        print("[HATA] Resim okunamadı:", image_path)
        return None
    plt.imshow(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    plt.title("Orijinal Görüntü")
    plt.show()
    student_roi = img[STUDENT_BOX_Y : STUDENT_BOX_Y + STUDENT_BOX_H,
                      STUDENT_BOX_X : STUDENT_BOX_X + STUDENT_BOX_W].copy()
    name_roi = img[NAME_BOX_Y : NAME_BOX_Y + NAME_BOX_H,
                   NAME_BOX_X : NAME_BOX_X + NAME_BOX_W].copy()
    answer_roi = img[ANSWER_BOX_Y : ANSWER_BOX_Y + ANSWER_BOX_H,
                     ANSWER_BOX_X : ANSWER_BOX_X + ANSWER_BOX_W].copy()
    plt.imshow(cv2.cvtColor(student_roi, cv2.COLOR_BGR2RGB))
    plt.title("Öğrenci No ROI")
    plt.show()
    plt.imshow(cv2.cvtColor(name_roi, cv2.COLOR_BGR2RGB))
    plt.title("Ad Soyad ROI")
    plt.show()
    plt.imshow(cv2.cvtColor(answer_roi, cv2.COLOR_BGR2RGB))
    plt.title("Cevaplar ROI")
    plt.show()
    student_no = read_student_number(student_roi)
    name_surname = read_name_surname(name_roi)
    student_answers = read_answers(answer_roi, ANSWER_KEY)
    results = compare_with_answer_key(student_answers, ANSWER_KEY)
    return {
        "student_no": student_no,
        "name_surname": name_surname,
        "answers": student_answers,
        "score_detail": results,  
        "correct_count": sum(1 for r in results if r[3] is True),
        "total": len(ANSWER_KEY)
    }
def read_student_number(roi):
    h, w = roi.shape[:2]
    margin_top = int((STUDENT_BOX_H - STUDENT_ROW_COUNT * STUDENT_BUBBLE_H) / 2)
    margin_left = int((STUDENT_BOX_W - STUDENT_COL_COUNT * STUDENT_BUBBLE_W) / 2)
    student_no_digits = []
    for row_index in range(STUDENT_ROW_COUNT):
        best_col = None
        best_ratio = 0.0
        for col_index in range(STUDENT_COL_COUNT):
            x1 = margin_left + col_index * STUDENT_BUBBLE_W
            y1 = margin_top + row_index * STUDENT_BUBBLE_H
            bubble_crop = roi[y1 : y1 + STUDENT_BUBBLE_H,
                              x1 : x1 + STUDENT_BUBBLE_W]
            ratio = get_fill_ratio(bubble_crop, THRESH_VAL)
            if ratio > best_ratio:
                best_ratio = ratio
                best_col = col_index
        if best_ratio >= FILL_RATIO_THRESHOLD and best_col is not None:
            student_no_digits.append(str(row_index))
        else:
            student_no_digits.append("-")
    return "".join(student_no_digits)
def read_name_surname(roi):
    name_chars = []
    h, w = roi.shape[:2]
    margin_top = int((NAME_BOX_H - NAME_ROW_COUNT * NAME_ROW_H) / 2)
    margin_left = int((NAME_BOX_W - NAME_COL_COUNT * NAME_BUBBLE_W) / 2)
    for col_index in range(NAME_COL_COUNT):
        best_row = None
        best_ratio = 0.0
        for row_index in range(NAME_ROW_COUNT):
            x1 = margin_left + col_index * NAME_BUBBLE_W
            y1 = margin_top + row_index * NAME_ROW_H  
            bubble_crop = roi[y1 : y1 + NAME_BUBBLE_H,
                              x1 : x1 + NAME_BUBBLE_W]
            ratio = get_fill_ratio(bubble_crop, THRESH_VAL)
            if ratio > best_ratio:
                best_ratio = ratio
                best_row = row_index
        if best_row is not None and best_ratio >= FILL_RATIO_THRESHOLD:
            if best_row < 26:
                letter = chr(ord('A') + best_row)
            else:
                letter = "?"  
            name_chars.append(letter)
        else:
            name_chars.append(" ")
    return "".join(name_chars).strip()
def read_answers(roi, answer_key):
    answers = []
    h, w = roi.shape[:2]
    margin_top = int((ANSWER_BOX_H - ANSWER_ROW_COUNT * ANSWER_ROW_H) / 2)
    margin_left = int((ANSWER_BOX_W - ANSWER_COL_COUNT * ANSWER_BUBBLE_W) / 2)
    for row_index in range(ANSWER_ROW_COUNT):
        best_col = None
        best_ratio = 0.0
        for col_index in range(1, ANSWER_COL_COUNT):  
            x1 = margin_left + col_index * ANSWER_BUBBLE_W
            y1 = margin_top + row_index * ANSWER_ROW_H
            bubble_crop = roi[y1 : y1 + ANSWER_BUBBLE_H,
                              x1 : x1 + ANSWER_BUBBLE_W]
            ratio = get_fill_ratio(bubble_crop, THRESH_VAL)
            if ratio > best_ratio:
                best_ratio = ratio
                best_col = col_index
        if best_ratio >= FILL_RATIO_THRESHOLD and best_col is not None:
            ans_char = chr(ord('A') + (best_col - 1))
            answers.append(ans_char)
        else:
            answers.append("")
    return answers
def compare_with_answer_key(student_answers, answer_key):
    results = []
    for i, correct_ans in enumerate(answer_key):
        ogr_ans = student_answers[i] if i < len(student_answers) else ""
        is_correct = (ogr_ans == correct_ans)
        results.append((i+1, ogr_ans, correct_ans, is_correct))
    return results
def get_fill_ratio(bubble_img, threshold_val=130):
    if bubble_img is None or bubble_img.size == 0:
        return 0.0
    gray = cv2.cvtColor(bubble_img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    _, th = cv2.threshold(blurred, threshold_val, 255, cv2.THRESH_BINARY_INV)
    kernel = np.ones((2, 2), np.uint8)
    th = cv2.morphologyEx(th, cv2.MORPH_CLOSE, kernel, iterations=1)
    black_pixels = np.count_nonzero(th)
    total_pixels = th.size
    return black_pixels / float(total_pixels)
if __name__ == "__main__":
    image_path = "C:/xmlayikla/1Optik Form.jpg"  
    result = read_omr_form(image_path)
    if result is None:
        print("[HATA] Okuma başarısız.")
    else:
        print("Öğrenci No  :", result["student_no"])
        print("Ad Soyad    :", result["name_surname"])
        print("-"*50)
        for (q_no, ogr_cvp, dogr_cvp, dogru_mu) in result["score_detail"]:
            durum_str = "Doğru" if dogru_mu else "Yanlış"
            print(f"Soru {q_no:02d}: Öğrenci Cevabı = {ogr_cvp or '-'} | "
                  f"Doğru Cevap = {dogr_cvp} | {durum_str}")
        print("-"*50)
        print(f"Toplam Doğru: {result['correct_count']} / {result['total']}")

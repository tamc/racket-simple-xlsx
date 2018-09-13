#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")
(require rackunit "../sheet.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-font-style"

    (let ([xlsx (new xlsx%)])

      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data 
       '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))])

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (fontColor . "red") (fontSize . 10) ))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 4))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (fontSize . 20) (fontColor . "00FF00")))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 8))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (fontSize . 10) (fontColor . "red")))
        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 12))
        
        (send xlsx burn-styles!)

        (let* ([xlsx_style (get-field style xlsx)]
               [cell_to_style_index_hash (data-sheet-cell_to_style_index_hash sheet)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [font_code_to_font_index_hash (xlsx-style-font_code_to_font_index_hash xlsx_style)]
               [font_list (xlsx-style-font_list xlsx_style)]
               [style1 (list-ref style_list 0)]
               [style2 (list-ref style_list 1)]
              )

          (check-equal? (hash-count cell_to_style_index_hash) 12)
          (check-equal? (length style_list) 2)
          (check-equal? (hash-count font_code_to_font_index_hash) 2)
          (check-equal? (length font_list) 2)
          
          (check-equal? (hash-ref cell_to_style_index_hash "A1") 1)
          (check-equal? (hash-ref cell_to_style_index_hash "A4") 1)
          (check-equal? (hash-ref cell_to_style_index_hash "B1") 2)
          (check-equal? (hash-ref cell_to_style_index_hash "B4") 2)
          (check-equal? (hash-ref cell_to_style_index_hash "C1") 1)
          (check-equal? (hash-ref cell_to_style_index_hash "C2") 1)

          (check-equal? (list-ref font_list (sub1 (hash-ref style1 'font))) (make-hash '((fontSize . 10) (fontColor . "red"))))
          (check-equal? (list-ref font_list (sub1 (hash-ref style2 'font))) (make-hash '((fontSize . 20) (fontColor . "00FF00"))))
          )

        )))))

(run-tests test-xlsx)

#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")
(require rackunit "../sheet.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-mix-style"

    (let ([xlsx (new xlsx%)])

      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data 
       '((1 2 3 4 5 6) (1 2 3 4 5 6) (1 2 3 4 5 6) (1 2 3 4 5 6) (1 2 3 4 5 6) (1 2 3 4 5 6)))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))])

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-B3" #:style '( (backgroundColor . "red") ))
        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B2-C4" #:style '( (backgroundColor . "blue") ))
        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C3-E4" #:style '( (fontSize . 10) ))
        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "D4-F5" #:style '( (fontSize . 5) ))

        (let* ([cell_to_style_code_hash (data-sheet-cell_to_style_code_hash sheet)]
               [style_code_to_style_hash (data-sheet-style_code_to_style_hash sheet)])

          (check-equal? (hash-count cell_to_style_code_hash) 25)
          (check-equal? (hash-count style_code_to_style_hash) 6))

        (send xlsx write-data-sheet-style! #:sheet_name "测试1")

        (let* ([xlsx_style (get-field style xlsx)]
               [cell_to_style_index_hash (data-sheet-cell_to_style_index_hash sheet)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [fill_list (xlsx-style-fill_list xlsx_style)]
               [font_list (xlsx-style-font_list xlsx_style)]
              )

          (check-equal? (hash-count cell_to_style_index_hash) 25)
          (check-equal? (length style_list) 6)
          (check-equal? (length fill_list) 2)
          (check-equal? (length font_list) 2)
          
          (check-equal? (hash-ref cell_to_style_index_hash "A1") 1)
          (check-equal? (hash-ref cell_to_style_index_hash "C4") 2)
          (check-equal? (hash-ref cell_to_style_index_hash "B2") 3)
          (check-equal? (hash-ref cell_to_style_index_hash "B4") 2)
          (check-equal? (hash-ref cell_to_style_index_hash "C1") 1)
          (check-equal? (hash-ref cell_to_style_index_hash "C2") 1)

          (check-equal? (list-ref fill_list (sub1 (hash-ref style1 'fill))) (make-hash '((fgColor . "red"))))
          (check-equal? (list-ref fill_list (sub1 (hash-ref style2 'fill))) (make-hash '((fgColor . "blue"))))
          )

        (let ([style_map (send xlsx get-cell-to-style-index-map "测试1")])
          (check-equal? (hash-count style_map) 12)
          (check-equal? (hash-ref style_map "A1") 1)
          (check-equal? (hash-ref style_map "B2") 2)
          (check-equal? (hash-ref style_map "C2") 1)
          )

        )))))

(run-tests test-xlsx)

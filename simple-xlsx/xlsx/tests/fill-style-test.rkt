#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")
(require rackunit "../sheet.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-fill-style"

    (let ([xlsx (new xlsx%)])

      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data 
       '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))])

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (backgroundColor . "red") ))
        (let* ([cell_to_style_code_hash (data-sheet-cell_to_style_code_hash sheet)]
               [style_code_to_style_hash (xlsx-style-style_code_to_style_hash xlsx_style)])

          (check-equal? (hash-count cell_to_style_code_hash) 4)
          (check-equal? (hash-count style_cell_to_style_code_hash) 1))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (backgroundColor . "blue") ))
        (let* ([cell_to_style_code_hash (data-sheet-cell_to_style_code_hash sheet)]
               [style_cell_to_style_code_hash (xlsx-style-style_cell_to_style_code_hash xlsx_style)])

          (check-equal? (hash-count cell_to_style_code_hash) 8)
          (check-equal? (hash-count style_cell_to_style_code_hash) 2))

        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (backgroundColor . "red") ))
        (let* ([cell_to_style_code_hash (data-sheet-cell_to_style_code_hash sheet)]
               [style_cell_to_style_code_hash (xlsx-style-style_cell_to_style_code_hash xlsx_style)])

          (check-equal? (hash-count cell_to_style_code_hash) 12)
          (check-equal? (hash-count style_code_to_style_hash) 2))

        (send xlsx write-data-sheet-style! #:sheet_name "测试1")

        (let* ([xlsx_style (get-field style xlsx)]
               [cell_to_style_index_hash (data-sheet-cell_to_style_index_hash sheet)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [fill_code_to_fill_index_hash (xlsx-style-fill_code_to_fill_index_hash xlsx_style)]
               [fill_list (xlsx-style-fill_list xlsx_style)]
              )

          (check-equal? (hash-count cell_to_style_index_hash) 12)
          (check-equal? (length style_list) 2)
          (check-equal? (hash-count fill_code_to_fill_index_hash) 2)
          (check-equal? (length fill_list) 2)
          
          (check-equal? (hash-ref cell_to_style_index_hash "A1") 1)
          (check-equal? (hash-ref cell_to_style_index_hash "A4") 1)
          (check-equal? (hash-ref cell_to_style_index_hash "B1") 2)
          (check-equal? (hash-ref cell_to_style_index_hash "B4") 2)
          (check-equal? (hash-ref cell_to_style_index_hash "C1") 1)
          (check-equal? (hash-ref cell_to_style_index_hash "C2") 1)
          )

        (let ([style_map (send xlsx get-cell-to-style-index-map "测试1")])
          (check-equal? (hash-count style_map) 12)
          (check-equal? (hash-ref style_map "A1") 1)
          (check-equal? (hash-ref style_map "B2") 2)
          (check-equal? (hash-ref style_map "C2") 1)
          )

        )))))

(run-tests test-xlsx)

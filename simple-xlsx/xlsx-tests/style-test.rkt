#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-set-data-sheet-cell-style-and-get-style-hash"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data 
       '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))

      (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (fgColor . "red") ))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))]
             [range_to_code_hash (data-sheet-range_to_style_code_hash sheet)]
             [code_to_style_hash (data-sheet-style_code_to_style_hash sheet)]
             [style_list (data-sheet-style_list sheet)]
             [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)])

       (check-equal? (hash-count range_to_code_hash) 1)
       (check-equal? (hash-count code_to_style_hash) 1)
       (check-equal? (length style_list) 1)
       (check-equal? (hash-count range_to_index_hash) 1)

       (let* ([style_index (hash-ref range_to_index_hash "A1-A4")]
              [style (list-ref style_list style_index)])

        (check-equal? style_index 0)
        (check-equal? (hash-count style) 1)
        (check-equal? (hash-ref style (quote fgColor)) "red")))

      (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (fgColor . "blue") ))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))]
             [range_to_code_hash (data-sheet-range_to_style_code_hash sheet)]
             [code_to_style_hash (data-sheet-style_code_to_style_hash sheet)]
             [style_list (data-sheet-style_list sheet)]
             [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)])

       (check-equal? (hash-count range_to_code_hash) 2)
       (check-equal? (hash-count code_to_style_hash) 2)
       (check-equal? (length style_list) 2)
       (check-equal? (hash-count range_to_index_hash) 2)

       (let* ([style_index (hash-ref range_to_index_hash "B1-B4")]
              [style (list-ref style_list style_index)])
         
        (check-equal? style_index 1)
        (check-equal? (hash-count style) 1)
        (check-equal? (hash-ref style 'fgColor) "blue")))

      (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (fgColor . "red") ))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))]
             [range_to_code_hash (data-sheet-range_to_style_code_hash sheet)]
             [code_to_style_hash (data-sheet-style_code_to_style_hash sheet)]
             [style_list (data-sheet-style_list sheet)]
             [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)])

       (check-equal? (hash-count range_to_code_hash) 3)
       (check-equal? (hash-count code_to_style_hash) 2)
       (check-equal? (length style_list) 2)
       (check-equal? (hash-count range_to_index_hash) 3)

       (let* ([style_index (hash-ref range_to_index_hash "C1-C4")]
              [style (list-ref style_list style_index)])

        (check-equal? style_index 0)
        (check-equal? (hash-count style) 1)
        (check-equal? (hash-ref style (quote fgColor)) "red")))
      
      (let ([style_map (send xlsx get-range-to-style-index-map "测试1")])
        (check-equal? (hash-ref style_map "C1-C4") 0))

      ))))

(run-tests test-xlsx)

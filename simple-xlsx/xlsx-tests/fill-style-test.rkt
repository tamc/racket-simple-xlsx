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

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))])

        (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (fgColor . "red") ))

        (let* ([xlsx_style (get-field style xlsx)]
               [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
               [code_to_style_hash (xlsx-style-style_code_to_style_index_hash xlsx_style)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [fill_code_to_fill_index_hash (xlsx-style-fill_code_to_fill_index_hash xlsx_style)]
               [fill_list (xlsx-style-fill_list xlsx_style)]
              )

          (check-equal? (hash-count code_to_style_hash) 1)
          (check-equal? (length style_list) 1)
          (check-equal? (hash-count fill_code_to_fill_index_hash) 1)
          (check-equal? (length fill_list) 1)
          (check-equal? (hash-count range_to_index_hash) 1)

          (let* ([style_index (hash-ref range_to_index_hash "A1-A4")]
                 [style (list-ref style_list style_index)]
                 [fill_index (hash-ref style 'fill)]
                 [fill (list-ref fill_list fill_index)])

            (check-equal? style_index 0)
            (check-equal? fill_index 0)
            (check-equal? (hash-count style) 1)
            (check-equal? (hash-ref style 'fill) 0)

            (check-equal? (hash-count fill) 1)
            (check-equal? (hash-ref fill 'fgColor) "red")
            ))

        (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (fgColor . "blue") ))

        (let* ([xlsx_style (get-field style xlsx)]
               [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
               [code_to_style_hash (xlsx-style-style_code_to_style_index_hash xlsx_style)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [fill_code_to_fill_index_hash (xlsx-style-fill_code_to_fill_index_hash xlsx_style)]
               [fill_list (xlsx-style-fill_list xlsx_style)]
              )

          (check-equal? (hash-count code_to_style_hash) 2)
          (check-equal? (length style_list) 2)
          (check-equal? (hash-count fill_code_to_fill_index_hash) 2)
          (check-equal? (length fill_list) 2)
          (check-equal? (hash-count range_to_index_hash) 2)

          (let* ([style_index (hash-ref range_to_index_hash "B1-B4")]
                 [style (list-ref style_list style_index)]
                 [fill_index (hash-ref style 'fill)]
                 [fill (list-ref fill_list fill_index)])

            (check-equal? style_index 1)
            (check-equal? fill_index 1)
            (check-equal? (hash-count style) 1)
            (check-equal? (hash-ref style 'fill) 1)

            (check-equal? (hash-count fill) 1)
            (check-equal? (hash-ref fill 'fgColor) "blue")
            ))

        (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (fgColor . "red") ))

        (let* ([xlsx_style (get-field style xlsx)]
               [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
               [code_to_style_hash (xlsx-style-style_code_to_style_index_hash xlsx_style)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [fill_code_to_fill_index_hash (xlsx-style-fill_code_to_fill_index_hash xlsx_style)]
               [fill_list (xlsx-style-fill_list xlsx_style)]
              )

          (check-equal? (hash-count code_to_style_hash) 2)
          (check-equal? (length style_list) 2)
          (check-equal? (hash-count fill_code_to_fill_index_hash) 2)
          (check-equal? (length fill_list) 2)
          (check-equal? (hash-count range_to_index_hash) 3)

          (let* ([style_index (hash-ref range_to_index_hash "C1-C4")]
                 [style (list-ref style_list style_index)]
                 [fill_index (hash-ref style 'fill)]
                 [fill (list-ref fill_list fill_index)])

            (check-equal? style_index 0)
            (check-equal? fill_index 0)
            (check-equal? (hash-count style) 1)
            (check-equal? (hash-ref style 'fill) 0)

            (check-equal? (hash-count fill) 1)
            (check-equal? (hash-ref fill 'fgColor) "red")
            ))

        (let ([style_map (send xlsx get-cell-to-style-index-map "测试1")])
          (check-equal? (hash-count style_map) 12)
          (check-equal? (hash-ref style_map "A1") 0)
          (check-equal? (hash-ref style_map "B2") 1)
          (check-equal? (hash-ref style_map "C2") 0)
          )

        )))))

(run-tests test-xlsx)

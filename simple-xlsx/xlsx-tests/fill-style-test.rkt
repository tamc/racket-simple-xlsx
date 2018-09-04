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

        (let ([code_to_style_hash (data-sheet-style_code_to_style_index_hash sheet)]
              [style_list (data-sheet-style_list sheet)]
              [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
              [fill_code_to_fill_index_hash (data-sheet-fill_code_to_fill_index_hash sheet)]
              [fill_list (data-sheet-fill_list sheet)]
              )

          (check-equal? 1 (hash-count code_to_style_hash))
          (check-equal? 1 (length style_list))
          (check-equal? 1 (hash-count fill_code_to_fill_index_hash))
          (check-equal? 1 (length fill_list))
          (check-equal? 1 (hash-count range_to_index_hash))

          (let* ([style_index (hash-ref range_to_index_hash "A1-A4")]
                 [style (list-ref style_list style_index)]
                 [fill_index (hash-ref style 'fill)]
                 [fill (list-ref fill_list fill_index)])

            (check-equal? 0 style_index)
            (check-equal? 0 fill_index)
            (check-equal? 1 (hash-count style))
            (check-equal? 0 (hash-ref style 'fill))

            (check-equal? 1 (hash-count fill))
            (check-equal? "red" (hash-ref fill 'fgColor))
            ))

        (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (fgColor . "blue") ))

        (let ([code_to_style_hash (data-sheet-style_code_to_style_index_hash sheet)]
              [style_list (data-sheet-style_list sheet)]
              [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
              [fill_code_to_fill_index_hash (data-sheet-fill_code_to_fill_index_hash sheet)]
              [fill_list (data-sheet-fill_list sheet)]
              )

          (check-equal? 2 (hash-count code_to_style_hash))
          (check-equal? 2 (length style_list))
          (check-equal? 2 (hash-count fill_code_to_fill_index_hash))
          (check-equal? 2 (length fill_list))
          (check-equal? 2 (hash-count range_to_index_hash))

          (let* ([style_index (hash-ref range_to_index_hash "B1-B4")]
                 [style (list-ref style_list style_index)]
                 [fill_index (hash-ref style 'fill)]
                 [fill (list-ref fill_list fill_index)])

            (check-equal? 1 style_index)
            (check-equal? 1 fill_index)
            (check-equal? 1 (hash-count style))
            (check-equal? 1 (hash-ref style 'fill))

            (check-equal? 1 (hash-count fill))
            (check-equal? "blue" (hash-ref fill 'fgColor))
            ))

        (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (fgColor . "red") ))

        (let ([code_to_style_hash (data-sheet-style_code_to_style_index_hash sheet)]
              [style_list (data-sheet-style_list sheet)]
              [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
              [fill_code_to_fill_index_hash (data-sheet-fill_code_to_fill_index_hash sheet)]
              [fill_list (data-sheet-fill_list sheet)]
              )

          (check-equal? 2 (hash-count code_to_style_hash))
          (check-equal? 2 (length style_list))
          (check-equal? 2 (hash-count fill_code_to_fill_index_hash))
          (check-equal? 2 (length fill_list))
          (check-equal? 3 (hash-count range_to_index_hash))

          (let* ([style_index (hash-ref range_to_index_hash "C1-C4")]
                 [style (list-ref style_list style_index)]
                 [fill_index (hash-ref style 'fill)]
                 [fill (list-ref fill_list fill_index)])

            (check-equal? 0 style_index)
            (check-equal? 0 fill_index)
            (check-equal? 1 (hash-count style))
            (check-equal? 0 (hash-ref style 'fill))

            (check-equal? 1 (hash-count fill))
            (check-equal? "red" (hash-ref fill 'fgColor))
            ))

        (let ([style_map (send xlsx get-cell-to-style-index-map "测试1")])
          (check-equal? (hash-count style_map) 12)
          (check-equal? (hash-ref style_map "A1") 0)
          (check-equal? (hash-ref style_map "B2") 1)
          (check-equal? (hash-ref style_map "C2") 0)
          )

        )))))

(run-tests test-xlsx)

#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")
(require rackunit "../sheet.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-set-data-sheet-cell-style-and-get-style-hash"

    (let ([xlsx (new xlsx%)])

      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data 
       '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))

      (let* ([sheet (sheet-content (send xlsx get-sheet-by-name "测试1"))])

        (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (fontSize . 20) ))

        (let* ([xlsx_style (get-field style xlsx)]
               [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
               [code_to_style_hash (xlsx-style-style_code_to_style_index_hash xlsx_style)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [font_code_to_font_index_hash (xlsx-style-font_code_to_font_index_hash xlsx_style)]
               [font_list (xlsx-style-font_list xlsx_style)]
              )

          (check-equal? (hash-count code_to_style_hash) 1)
          (check-equal? (length style_list) 1)
          (check-equal? (hash-count font_code_to_font_index_hash) 1)
          (check-equal? (length font_list) 1)
          (check-equal? (hash-count range_to_index_hash) 1)

          (let* ([style_index (hash-ref range_to_index_hash "A1-A4")]
                 [style (list-ref style_list (sub1 style_index))]
                 [font_index (hash-ref style 'font)]
                 [font (list-ref font_list (sub1 font_index))])

            (check-equal? style_index 1)
            (check-equal? font_index 1)
            (check-equal? (hash-count style) 1)
            (check-equal? (hash-ref style 'font) 1)

            (check-equal? (hash-count font) 1)
            (check-equal? (hash-ref font 'fontSize) 20)
            ))

        (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (fontSize . 30) ))

        (let* ([xlsx_style (get-field style xlsx)]
               [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
               [code_to_style_hash (xlsx-style-style_code_to_style_index_hash xlsx_style)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [font_code_to_font_index_hash (xlsx-style-font_code_to_font_index_hash xlsx_style)]
               [font_list (xlsx-style-font_list xlsx_style)]
              )

          (check-equal? (hash-count code_to_style_hash) 2)
          (check-equal? (length style_list) 2)
          (check-equal? (hash-count font_code_to_font_index_hash) 2)
          (check-equal? (length font_list) 2)
          (check-equal? (hash-count range_to_index_hash) 2)

          (let* ([style_index (hash-ref range_to_index_hash "B1-B4")]
                 [style (list-ref style_list (sub1 style_index))]
                 [font_index (hash-ref style 'font)]
                 [font (list-ref font_list (sub1 font_index))])

            (check-equal? style_index 2)
            (check-equal? font_index 2)
            (check-equal? (hash-count style) 1)
            (check-equal? (hash-ref style 'font) 2)

            (check-equal? (hash-count font) 1)
            (check-equal? (hash-ref font 'fontSize) 30)
            ))

        (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (fontSize . 20) ))

        (let* ([xlsx_style (get-field style xlsx)]
               [range_to_index_hash (data-sheet-range_to_style_index_hash sheet)]
               [code_to_style_hash (xlsx-style-style_code_to_style_index_hash xlsx_style)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [font_code_to_font_index_hash (xlsx-style-font_code_to_font_index_hash xlsx_style)]
               [font_list (xlsx-style-font_list xlsx_style)]
              )

          (check-equal? (hash-count code_to_style_hash) 2)
          (check-equal? (length style_list) 2)
          (check-equal? (hash-count font_code_to_font_index_hash) 2)
          (check-equal? (length font_list) 2)
          (check-equal? (hash-count range_to_index_hash) 3)

          (let* ([style_index (hash-ref range_to_index_hash "C1-C4")]
                 [style (list-ref style_list (sub1 style_index))]
                 [font_index (hash-ref style 'font)]
                 [font (list-ref font_list (sub1 font_index))])

            (check-equal? style_index 1)
            (check-equal? font_index 1)
            (check-equal? (hash-count style) 1)
            (check-equal? (hash-ref style 'font) 1)

            (check-equal? (hash-count font) 1)
            (check-equal? (hash-ref font 'fontSize) 20)
            ))
        
        )))))

(run-tests test-xlsx)

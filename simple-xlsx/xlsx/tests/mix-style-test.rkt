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
        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B2-D3" #:style '( (backgroundColor . "blue") ))
        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C3-E4" #:style '( (fontSize . 10) ))
        (send xlsx add-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "E4-F6" #:style '( (fontSize . 5) ))

        (let* ([cell_to_origin_style_hash (data-sheet-cell_to_origin_style_hash sheet)])
          (check-equal? (hash-count cell_to_origin_style_hash) 19))

        (send xlsx write-data-sheet-style! #:sheet_name "测试1")

        (let* ([xlsx_style (get-field style xlsx)]
               [cell_to_style_index_hash (data-sheet-cell_to_style_index_hash sheet)]
               [style_list (xlsx-style-style_list xlsx_style)]
               [fill_list (xlsx-style-fill_list xlsx_style)]
               [font_list (xlsx-style-font_list xlsx_style)]
              )

          (check-equal? (hash-count cell_to_style_index_hash) 19)
          (check-equal? (length style_list) 5)
          (check-equal? (length fill_list) 2)
          (check-equal? (length font_list) 2)

          (check-equal? (hash-count (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "A1")))) 1)
          (check-equal? 
           (list-ref fill_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "A1"))) 'fill)))
           (make-hash '((fgColor . "red"))))

          (check-equal? (hash-count (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "B2")))) 1)
          (check-equal? 
           (list-ref fill_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "B2"))) 'fill)))
           (make-hash '((fgColor . "blue"))))

          (check-equal? (hash-count (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "C3")))) 2)
          (check-equal? 
           (list-ref fill_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "C3"))) 'fill)))
           (make-hash '((fgColor . "blue"))))
          (check-equal? 
           (list-ref font_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "C3"))) 'font)))
           (make-hash '((fontSize . 10))))

          (check-equal? (hash-count (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "D4")))) 1)
          (check-equal? 
           (list-ref font_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "D4"))) 'font)))
           (make-hash '((fontSize . 10))))

          (check-equal? (hash-count (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "F6")))) 1)
          (check-equal? 
           (list-ref font_list (sub1 (hash-ref (list-ref style_list (sub1 (hash-ref cell_to_style_index_hash "F6"))) 'font)))
           (make-hash '((fontSize . 5))))
          )

        )))))

(run-tests test-xlsx)

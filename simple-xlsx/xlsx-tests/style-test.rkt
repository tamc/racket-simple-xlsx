#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-set-data-sheet-cell-style-and-get-style-hash"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "测试1" #:sheet_data '((1 2 "chenxiao") (3 4 "xiaomin") (5 6 "chenxiao") (1 "xx" "simmone")))

      (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "A1-A4" #:style '( (fgColor . "red") ))

      (let* ([style_hash (send xlsx get-style-hash)]
             [style_list (send xlsx get-style-list)]
             [style_index (hash-ref style_hash "A1-A4")]
             [style (hash-ref style_list style_index)])

        (check-equal? (hash-count style_hash) 1)
        (check-equal? (length style_list) 1)
        (check-true (hash? style))
        (check-equal? (hash-count style) 1)
        (check-equal? (hash-ref style 'fgColor) "red"))

      (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "B1-B4" #:style '( (fgColor . "blue") ))

      (let* ([style_hash (send xlsx get-style-hash)]
             [style_list (send xlsx get-style-list)]
             [style_index (hash-ref style_hash "B1-B4")]
             [style (hash-ref style_list style_index)])

        (check-equal? (hash-count style_hash) 2)
        (check-equal? (length style_list) 2)
        (check-equal? (hash-ref style 'fgColor) "blue"))

      (send xlsx set-data-sheet-cell-style! #:sheet_name "测试1" #:cell_range "C1-C4" #:style '( (fgColor . "blue") ))

      (let* ([style_hash (send xlsx get-style-hash)]
             [style_list (send xlsx get-style-list)]
             [style_index (hash-ref style_hash "C1-C4")]
             [style (hash-ref style_list style_index)])

        (check-equal? (hash-count style_hash) 2)
        (check-equal? (length style_list) 2)
        (check-equal? (hash-ref style 'fgColor) "blue"))

      ))

   ))

(run-tests test-xlsx)

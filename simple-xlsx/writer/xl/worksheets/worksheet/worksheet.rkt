#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../../../lib/lib.rkt")
(require "../../../../xlsx/xlsx.rkt")
(require "../../../../xlsx/sheet.rkt")

(provide (contract-out
          [write-header (-> string?)]
          [write-data-sheet (-> string? (is-a?/c xlsx%) string?)]
          [write-data-sheet-file (-> path-string? (is-a?/c xlsx%) void?)]
          [get-col-width-map (-> (listof list?) hash?)]
          ))

(define S string-append)

(define (write-header) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<worksheet
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
})

(define (write-dimension dimension) @S{
  <dimension ref="@|dimension|"/>
})

(define (write-sheet-views is_active active_cell) @S{
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="A1" sqref="A1"/>
    </sheetView>
  </sheetViews>

  <dimension ref="@|dimension|"/>
})

(define (write-data-sheet sheet_name xlsx)
  (let ([string_index_map (send xlsx get-string-index-map)]
        [cell_to_style_index_hash (send xlsx get-cell-to-style-index-map sheet_name)])

    (with-output-to-string
      (lambda ()
        ;; only data sheet
        (let* ([sheet (send xlsx get-sheet-by-name sheet_name)]
               [data_sheet (sheet-content sheet)]
               [rows (data-sheet-rows data_sheet)]
               [col_width_map (get-col-width-map rows)]
               [col_count (length (car rows))]
               [span_str (string-append "1:" (number->string col_count))]
               [width_hash (data-sheet-width_hash data_sheet)]
               [dimension (if (= (length rows) 0) "A1" (string-append "A1:" (get-dimension rows)))]
               [is_active (if (= (sheet-seq sheet) 1) "tabSelected=\"1\"" "")]
               [active_cell (if (null? data_sheet) "" "<selection activeCell=\"A1\" sqref=\"A1\"/>")])
               
          (printf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
          (printf "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><dimension ref=\"~a\"/><sheetViews><sheetView ~a workbookViewId=\"0\">~a</sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"13.5\"/>" dimension is_active active_cell)

          (hash-for-each
           width_hash
           (lambda (col_range width)
             (let* ([items (regexp-match* #rx"([A-Z]+)" col_range)]
                    [start_index (abc->number (first items))]
                    [end_index (abc->number (second items))])
               (let loop ([loop_index start_index])
                 (when (<= loop_index end_index)
                       (hash-set! col_width_map loop_index width)
                       (loop (add1 loop_index)))))))

          (printf "<cols>")

          (for-each
           (lambda (col)
             (printf "<col min=\"~a\" max=\"~a\" width=\"~a\"/>" (car col) (car col) (cdr col)))
           (sort (hash->list col_width_map) < #:key car))
           
          (printf "</cols>")
          
          (printf "<sheetData>")
          
          (let loop-row ([loop_rows (data-sheet-rows data_sheet)]
                         [row_seq 1])
            (when (not (null? loop_rows))
                  (printf "<row r=\"~a\" spans=\"~a\">" row_seq span_str)
                  (let ([item_list (car loop_rows)])
                    (when (not (null? item_list))
                          (let loop-col ([loop_cols item_list]
                                         [col_seq 1])
                            (when (not (null? loop_cols))
                                  (let* ([cell (car loop_cols)]
                                         [dimension (string-append (number->abc col_seq) (number->string row_seq))]
                                         [style_index (hash-ref cell_to_style_index_hash dimension #f)]
                                         [style
                                          (if style_index (string-append " s=\"" (number->string style_index) "\"") "")])
                                    (cond
                                     [(string? cell)
                                      (printf "<c r=\"~a\"~a t=\"s\"><v>~a</v></c>" 
                                              dimension style (hash-ref string_index_map cell))]
                                     [(exact-integer? cell)
                                      (printf "<c r=\"~a\"~a><v>~a</v></c>" 
                                              dimension style (number->string (inexact->exact cell)))]
                                     [(number? cell)
                                      (printf "<c r=\"~a\"~a><v>~a</v></c>" 
                                              dimension style (number->string (exact->inexact cell)))]
                                     [else
                                      (printf "<c r=\"~a\"><v>0</v></c>" dimension)]))
                                  (loop-col (cdr loop_cols) (add1 col_seq))))))
                  (printf "</row>")
                  (loop-row (cdr loop_rows) (add1 row_seq)))))

        (printf "</sheetData><phoneticPr fontId=\"1\" type=\"noConversion\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"200\" verticalDpi=\"200\" r:id=\"rId1\"/></worksheet>")))))

(define (write-data-sheet-file dir xlsx)
  (make-directory* dir)

  (let loop ([loop_list (get-field sheets xlsx)])
    (when (not (null? loop_list))
          (when (eq? (sheet-type (car loop_list)) 'data)
                (with-output-to-file (build-path dir (string-append "sheet" (number->string (sheet-typeSeq (car loop_list))) ".xml"))
                  #:exists 'replace
                  (lambda ()
                    (printf "~a" (write-data-sheet (sheet-name (car loop_list)) xlsx)))))
          (loop (cdr loop_list)))))

(define (get-col-width-map rows)
  (let ([col_width_map (make-hash)])
    (let loop-row ([loop_rows rows])
      (when (not (null? loop_rows))
            (let loop-col ([cols (car loop_rows)]
                           [index 1])
              (when (not (null? cols))
                    (let* ([str (if (number? (car cols)) (number->string (car cols)) (car cols))]
                           [str_len (string-length str)]
                           [bytes_len (bytes-length (string->bytes/utf-8 str))]
                           [result_len (+ 2 (- bytes_len (floor (/ (- bytes_len str_len) 2))))])
                      (when (< (hash-ref col_width_map index 0) result_len)
                            (hash-set! col_width_map index result_len)))
                    (loop-col (cdr cols) (add1 index))))
            (loop-row (cdr loop_rows))))
    col_width_map))

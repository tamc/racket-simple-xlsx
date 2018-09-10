#lang racket

(require "../lib/lib.rkt")
(require "xlsx-lib.rkt")
(require "sheet.rkt")

(provide (contract-out
          [xlsx% class?]
          [struct xlsx-style
                  (
                   (style_code_to_style_index_hash hash?)
                   (style_list list?)
                   (fill_code_to_fill_index_hash hash?)
                   (fill_list list?)
                   (font_code_to_font_index_hash hash?)
                   (font_list list?)
                   )]
          ))

(struct xlsx-style (
                    [style_code_to_style_index_hash #:mutable]
                    [style_list #:mutable]
                    [fill_code_to_fill_index_hash #:mutable]
                    [fill_list #:mutable]
                    [font_code_to_font_index_hash #:mutable]
                    [font_list #:mutable]
                    ))

(define xlsx%
  (class object%
         (super-new)
         
         (field 
          [sheets '()]
          [sheet_name_map (make-hash)]
          [string_item_map (make-hash)]
          [style (xlsx-style (make-hash) '() (make-hash) '() (make-hash) '())]
          )
         
         (define/public (add-data-sheet #:sheet_name sheet_name #:sheet_data sheet_data)
           (when (check-data-list sheet_data)
                 (if (not (hash-has-key? sheet_name_map sheet_name))
                     (let* ([sheet_length (length sheets)]
                            [seq (add1 sheet_length)]
                            [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'data)) sheets)))])
                       
                       (let loop ([loop_list sheet_data])
                         (when (not (null? loop_list))
                               (let inner-loop ([inner_loop_list (car loop_list)])
                                 (when (not (null? inner_loop_list))
                                       (when (string? (car inner_loop_list))
                                             (hash-set! string_item_map (car inner_loop_list) ""))
                                       (inner-loop (cdr inner_loop_list))))
                               (loop (cdr loop_list))))
                       
                       (set! sheets `(,@sheets
                                      ,(sheet
                                        sheet_name
                                        seq
                                        'data
                                        type_seq
                                        (data-sheet sheet_data (make-hash) (make-hash)))))
                       (hash-set! sheet_name_map sheet_name (sub1 seq)))
                     (error (format "duplicate sheet name[~a]" sheet_name)))))
         
         (define/public (get-sheet-by-name sheet_name)
           (if (hash-has-key? sheet_name_map sheet_name)
               (list-ref sheets (hash-ref sheet_name_map sheet_name))
               (error (format "no such sheet[~a]" sheet_name))))
         
         (define/public (get-range-data sheet_name range_str)
           (when (check-data-range-valid #:sheet_name sheet_name #:range_str range_str)
                 (let* ([data_sheet (get-sheet-by-name sheet_name)]
                        [col_range (regexp-match* #rx"([A-Z]+)" range_str)]
                        [start_col_name (first col_range)]
                        [col_start_index (sub1 (abc->number start_col_name))]
                        [end_col_name (second col_range)]
                        [col_end_index (sub1 (abc->number end_col_name))]
                        [row_range (regexp-match* #rx"([0-9]+)" range_str)]
                        [row_start_index (string->number (first row_range))]
                        [row_end_index (string->number (second row_range))]
                        [direction (if (string=? start_col_name end_col_name) 'vertical 'horizontal)])
                   (if (eq? direction 'vertical)
                       (let loop ([loop_list (data-sheet-rows (sheet-content data_sheet))]
                                  [row_count 1]
                                  [result_list '()])
                         (if (not (null? loop_list))
                             (if (and (>= row_count row_start_index) (<= row_count row_end_index))
                                 (loop (cdr loop_list) (add1 row_count) 
                                       (cons (list-ref (car loop_list) col_start_index) result_list))
                                 (loop (cdr loop_list) (add1 row_count) result_list))
                             (reverse result_list)))
                       (let loop ([loop_list (list-ref (data-sheet-rows (sheet-content data_sheet)) (sub1 row_start_index))]
                                  [col_index 0]
                                  [result_list '()])
                         (if (not (null? loop_list))
                             (if (and (>= col_index col_start_index) (<= col_index col_end_index))
                                 (loop (cdr loop_list) (add1 col_index) (cons (car loop_list) result_list))
                                 (loop (cdr loop_list) (add1 col_index) result_list))
                             (reverse result_list)))))))

         (define/public (add-chart-sheet #:sheet_name sheet_name #:chart_type [chart_type 'line] #:topic [topic ""] #:x_topic [x_topic ""])
           (if (not (hash-has-key? sheet_name_map sheet_name))
               (let* ([sheet_length (length sheets)]
                      [seq (add1 sheet_length)]
                      [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'chart)) sheets)))])
                 (set! sheets `(,@sheets
                                ,(sheet
                                  sheet_name
                                  seq
                                  'chart
                                  type_seq
                                  (chart-sheet chart_type topic x_topic (data-range "" "") '()))))
                 (hash-set! sheet_name_map sheet_name (sub1 seq)))
               (error (format "duplicate sheet name[~a]" sheet_name))))

         (define/public (check-data-range-valid #:sheet_name sheet_name #:range_str range_str)
           (when (check-range range_str)
                 (let* ([rows (data-sheet-rows (sheet-content (get-sheet-by-name sheet_name)))]
                        [first_row (first rows)]
                        [col_name (first (regexp-match* #rx"([A-Z]+)" range_str))]
                        [col_number (sub1 (abc->number col_name))]
                        [row_range (regexp-match* #rx"([0-9]+)" range_str)]
                        [row_start_index (string->number (first row_range))]
                        [row_end_index (string->number (second row_range))])
                   (cond
                    [(< (length first_row) (add1 col_number))
                     (error (format "no such column[~a]" col_name))]
                    [(> row_end_index (length rows))
                     (error (format "end index beyond data range[~a]" row_end_index))]
                    [else
                     #t]))))
                 
         (define/public (set-chart-x-data! #:sheet_name sheet_name #:data_sheet_name data_sheet_name #:data_range data_range)
           (when (check-data-range-valid #:sheet_name data_sheet_name #:range_str data_range)
                 (set-chart-sheet-x_data_range! (sheet-content (get-sheet-by-name sheet_name)) (data-range data_sheet_name data_range))))

         (define/public 
           (add-chart-serial! #:sheet_name sheet_name #:data_sheet_name data_sheet_name #:data_range data_range #:y_topic [y_topic ""])
           (when (check-data-range-valid #:sheet_name data_sheet_name #:range_str data_range)
                 (set-chart-sheet-y_data_range_list! (sheet-content (get-sheet-by-name sheet_name)) `(,@(chart-sheet-y_data_range_list (sheet-content (get-sheet-by-name sheet_name))) ,(data-serial y_topic (data-range data_sheet_name data_range))))))

         (define/public (sheet-ref sheet_seq)
           (list-ref sheets sheet_seq))
         
         (define/public (set-data-sheet-col-width! #:sheet_name sheet_name #:col_range col_range #:width width)
           (let ([converted_col_range (check-col-range col_range)])
                 (hash-set! (data-sheet-width_hash (sheet-content (get-sheet-by-name sheet_name))) converted_col_range width)))

         (define/public (get-string-item-list)
           (sort (hash-keys string_item_map) string<?))

         (define/public (get-string-index-map)
           (let ([string_index_map (make-hash)])
             (let loop ([loop_list (get-string-item-list)]
                        [index 0])
               (when (not (null? loop_list))
                     (hash-set! string_index_map (car loop_list) index)
                     (loop (cdr loop_list) (add1 index))))
             string_index_map))

         (define/public (write-data-sheet-cell-style! #:sheet_name sheet_name)
           (when (check-cell-range cell_range)
                 (let* ([sheet (sheet-content (get-sheet-by-name sheet_name))]
                        [range_to_style_index_hash (data-sheet-range_to_style_index_hash sheet)]
                        [style_code_to_style_index_hash (xlsx-style-style_code_to_style_index_hash style)]
                        [style_list (xlsx-style-style_list style)]
                        [style_hash (make-hash)]
                        [style_hash_code #f]
                        [fill_hash (make-hash)]
                        [fill_hash_code #f]
                        [fill_code_to_fill_index_hash (xlsx-style-fill_code_to_fill_index_hash style)]
                        [fill_list (xlsx-style-fill_list style)]
                        [font_hash (make-hash)]
                        [font_hash_code #f]
                        [font_code_to_font_index_hash (xlsx-style-font_code_to_font_index_hash style)]
                        [font_list (xlsx-style-font_list style)]
                        )
                   
                   (for-each
                    (lambda (style_pair)
                      (when (and
                             (pair? style_pair)
                             (symbol? (car style_pair)))
                            (cond
                             [(or
                               (eq? (car style_pair) 'backgroundColor)
                               )
                              (hash-set! fill_hash (car style_pair) (cdr style_pair))]
                             [(or
                               (eq? (car style_pair) 'fontSize)
                               )
                              (hash-set! font_hash (car style_pair) (cdr style_pair))]
                             )))
                    style_pair_list)
                   
                   (when (> (hash-count fill_hash) 0)
                         (set! fill_hash_code (equal-hash-code fill_hash))

                         (if (not (hash-has-key? fill_code_to_fill_index_hash fill_hash_code))
                             (begin
                               (hash-set! fill_code_to_fill_index_hash fill_hash_code (add1 (length fill_list)))
                               (set-xlsx-style-fill_list! style `(,@fill_list ,fill_hash))
                               (hash-set! style_hash 'fill (add1 (length fill_list))))
                             (hash-set! style_hash 'fill (hash-ref fill_code_to_fill_index_hash fill_hash_code))))

                   (when (> (hash-count font_hash) 0)
                         (set! font_hash_code (equal-hash-code font_hash))

                         (if (not (hash-has-key? font_code_to_font_index_hash font_hash_code))
                             (begin
                               (hash-set! font_code_to_font_index_hash font_hash_code (add1 (length font_list)))
                               (set-xlsx-style-font_list! style `(,@font_list ,font_hash))
                               (hash-set! style_hash 'font (add1 (length font_list))))
                             (hash-set! style_hash 'font (hash-ref font_code_to_font_index_hash font_hash_code))))
                   
                   (when (> (hash-count style_hash) 0)
                         (set! style_hash_code (equal-hash-code style_hash))
                         
                         (if (not (hash-has-key? style_code_to_style_index_hash style_hash_code))
                             (begin
                               (hash-set! style_code_to_style_index_hash style_hash_code (add1 (length style_list)))
                               (set-xlsx-style-style_list! style `(,@style_list ,style_hash))
                               (hash-set! range_to_style_index_hash cell_range (add1 (length style_list))))
                             (hash-set! range_to_style_index_hash cell_range (hash-ref style_code_to_style_index_hash style_hash_code))))
                   )))

         (define/public (get-cell-to-style-index-map sheet_name)
           (flat-range-hash
            (data-sheet-range_to_style_index_hash (sheet-content (get-sheet-by-name sheet_name)))))

         (define/public (get-style-list) (xlsx-style-style_list style))

         (define/public (get-fill-list) (xlsx-style-fill_list style))

         (define/public (get-font-list) (xlsx-style-font_list style))

         ))

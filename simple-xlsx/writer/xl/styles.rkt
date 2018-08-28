#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../lib/lib.rkt")
(require "../../xlsx.rkt")

(provide (contract-out
          [write-styles (-> list? string?)]
          [write-styles-file (-> path-string? (is-a?/c xlsx%) void?)]
          [write-header (-> string?)]
          [write-fonts (-> string?)]
          [write-fills (-> string?)]
          [write-borders (-> string?)]
          [write-cellStyleXfs (-> string?)]
          [write-cellXfs (-> string?)]
          [write-cellStyles (-> string?)]
          [write-dxfs (-> string?)]
          [write-footer (-> string?)]
          ))

(define S string-append)

(define (write-header) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
})

(define (write-fonts) @S{
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="宋体"/>
      <family val="2"/>
      <charset val="134"/>
      <scheme val="minor"/>
    </font>
  </fonts>
})

(define (write-fills) @S{
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0000"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="00FF00"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="0000FF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
})

(define (write-borders) @S{
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
})

(define (write-cellStyleXfs) @S{
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf>
  </cellStyleXfs>
})

(define (write-cellXfs) @S{
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"><alignment vertical="center"/></xf>
  </cellXfs>
})

(define (write-cellStyles) @S{
  <cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>
})

(define (write-dxfs) @S{
  <dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
})

(define (write-footer) @S{
</styleSheet>
})

(define (write-styles color_list) @S{
@|(write-header)|

@|(write-fonts)|

@|(write-fills)|

@|(write-borders)|

@|(write-cellStyleXfs)|

@|(write-cellXfs)|

@|(write-cellStyles)|

@|(write-dxfs)|

@|(write-footer)|
})

(define (write-styles-bak color_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font><font><sz val="9"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="@|(number->string (+ 2 (length color_list)))|"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>@|(with-output-to-string
  (lambda ()
    (for-each
     (lambda (style_rec)
       (printf "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"~a\"/><bgColor indexed=\"64\"/></patternFill></fill>" style_rec))
     color_list)))|</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="@|(number->string (add1 (length color_list)))|"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>@|(with-output-to-string
(lambda ()
  (let loop ([loop_list color_list]
             [index 2])
    (when (not (null? loop_list))
          (printf "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"~a\" borderId=\"0\" xfId=\"0\" applyFill=\"1\"><alignment vertical=\"center\"/></xf>" index)
          (loop (cdr loop_list) (add1 index))))))|</cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>
})

(define (write-styles-file dir xlsx)
  (make-directory* dir)

  (with-output-to-file (build-path dir "styles.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-styles (send xlsx get-color-list))))))

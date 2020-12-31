<?php

namespace Drupal\xls_serialization;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * The interface to implement by a field formatter plugin to style Excel cells.
 *
 * @see \Drupal\xls_serialization\ViewsExcelStyledFieldItemValue
 */
interface ViewsExcelStyledFieldCellFormatterInterface {

  /**
   * Returns a processed value and allows to apply styles for a cell.
   *
   * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet
   *   The worksheet.
   * @param string $coordinate
   *   The coordinates of a cell.
   * @param string $value
   *   The value for a cell.
   *
   * @return string
   *   The processed "$value".
   */
  public static function setStyle(Worksheet $sheet, string $coordinate, string $value): string;

}

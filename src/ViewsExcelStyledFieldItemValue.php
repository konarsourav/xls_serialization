<?php

namespace Drupal\xls_serialization;

use Drupal\Core\Field\FormatterInterface;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * The object to be used by a field formatter for applying styles to a cell.
 *
 * @example
 * The field formatter to be configured for Views Data Export. The "MyFormatter"
 * plugin must be selected as a formatter for a field in Views configuration.
 * @code
 * namespace Drupal\my_module\Plugin\Field\FieldFormatter;
 *
 * use Drupal\Core\Field\FormatterBase;
 * use Drupal\Core\Field\FieldItemListInterface;
 * use Drupal\xls_serialization\ViewsExcelStyledFieldCellFormatterInterface;
 * use Drupal\xls_serialization\ViewsExcelStyledFieldItemValue;
 * use PhpOffice\PhpSpreadsheet\Style\Color;
 * use PhpOffice\PhpSpreadsheet\Style\Fill;
 * use PhpOffice\PhpSpreadsheet\Style\Style;
 * use PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;
 * use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
 *
 * class MyFormatter extends FormatterBase implements ViewsExcelStyledFieldCellFormatterInterface {
 *   protected const COLORS = [
 *     'red' => Color::COLOR_RED,
 *     'orange' => 'FFFFA500',
 *     'green' => Color::COLOR_DARKGREEN,
 *     'purple' => [128, 0, 128],
 *     'blue' => Color::COLOR_BLUE,
 *     'grey' => 'FF808080',
 *   ];
 *
 *   public function viewElements(FieldItemListInterface $items, $langcode): array {
 *     $elements = [];
 *
 *     foreach ($items as $delta => $item) {
 *       $elements[$delta]['#markup'] = new ViewsExcelStyledFieldItemValue($item->value, static::class);
 *     }
 *
 *     return $elements;
 *   }
 *
 *   public static function setStyle(Worksheet $sheet, string $coordinate, string $value): string {
 *     if (!isset(static::COLORS[$value])) {
 *       // Return the name of a color since our mapping has no ARGB value
 *       // for the color name.
 *       return $value;
 *     }
 *
 *     is_array(static::COLORS[$value])
 *       ? static::drawCircle($sheet, $coordinate, ...static::COLORS[$value])
 *       : static::fillCell($sheet, $coordinate, new Color(static::COLORS[$value]));
 *
 *     // Return an empty string to make a table cell empty because we do
 *     // graphic instead of putting color name inside of a cell.
 *     return '';
 *   }
 *
 *   protected static function drawCircle(Worksheet $sheet, string $coordinate, int $red, int $green, int $blue): void {
 *     $image = imagecreate(12, 12);
 *     $drawing = new MemoryDrawing();
 *
 *     $drawing->setName($value);
 *     $drawing->setMimeType($drawing::MIMETYPE_DEFAULT);
 *     $drawing->setWorksheet($sheet);
 *     $drawing->setCoordinates($coordinate);
 *     $drawing->setImageResource($image);
 *     $drawing->setRenderingFunction($drawing::RENDERING_PNG);
 *
 *     // Width and height are available after calling "setImageResource()".
 *     $width = $drawing->getWidth();
 *     $height = $drawing->getHeight();
 *
 *     // Make image background transparent.
 *     imagefill($image, 0, 0, imagecolortransparent($image, imagecolorallocate($image, 250, 0, 0)));
 *     // Draw a colored circle.
 *     imagefilledellipse($image, $width / 2, $height / 2, $width, $height, imagecolorallocate($image, $red, $green, $blue));
 *   }
 *
 *   protected static function fillCell(Worksheet $sheet, string $coordinate, Color $color): void {
 *     $sheet
 *       ->getStyle($coordinate)
 *       ->getFill()
 *       ->setFillType(Fill::FILL_SOLID)
 *       ->setStartColor($color);
 *   }
 * }
 * @endcode
 *
 * @see \Drupal\xls_serialization\Encoder\Xls::setData()
 */
class ViewsExcelStyledFieldItemValue implements \Serializable {

  /**
   * The value to display within a document's cell.
   *
   * @var string|null
   */
  protected $value;

  /**
   * The callback function to alter cell's style.
   *
   * This function:
   * - receives two arguments:
   *   - \PhpOffice\PhpSpreadsheet\Style\Style $style
   *   - string $value
   * - must return second argument if a value doesn't need to be changed.
   *
   * @var string|null
   */
  protected $callback;

  /**
   * {@inheritdoc}
   */
  public function __construct(string $value, string $formatter) {
    if (!is_subclass_of($formatter, FormatterInterface::class, TRUE) || !is_subclass_of($formatter, ViewsExcelStyledFieldCellFormatterInterface::class, TRUE)) {
      throw new \InvalidArgumentException(sprintf('Argument #2 must be a field formatter that implements "%s" interface.', ViewsExcelStyledFieldCellFormatterInterface::class));
    }

    $this->value = $value;
    /* @see \Drupal\xls_serialization\ViewsExcelStyledFieldCellFormatterInterface::setStyle() */
    $this->callback = $formatter . '::setStyle';
  }

  /**
   * {@inheritdoc}
   */
  public function serialize(): string {
    return serialize([$this->value, $this->callback]);
  }

  /**
   * {@inheritdoc}
   */
  public function unserialize($serialized): void {
    [$this->value, $this->callback] = unserialize($serialized);
  }

  /**
   * Returns the serialized representation of this object.
   *
   * @return string
   *   The serialized representation of this object.
   */
  public function __toString(): string {
    return serialize($this);
  }

  /**
   * Returns the value for a document's cell.
   *
   * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet
   *   The worksheet.
   * @param string $coordinate
   *   The coordinates of a cell.
   *
   * @return string
   *   The original "$this->value" or something else if this needs
   *   to be changed.
   */
  public function setStyle(Worksheet $sheet, string $coordinate): string {
    return call_user_func($this->callback, $sheet, $coordinate, $this->value);
  }

}

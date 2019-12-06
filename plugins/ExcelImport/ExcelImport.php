<?php
/**
 * TestLink Open Source Project - http://testlink.sourceforge.net/
 * This script is distributed under the GNU General Public License 2 or later.
 *
 * @filesource  TLTest.php
 * @copyright   2005-2016, TestLink community
 * @link        http://www.testlink.org/
 *
 */

require_once(TL_ABS_PATH . '/lib/functions/tlPlugin.class.php');

/**
 * Sample Testlink Plugin class that registers itself with the system and provides 
 * UI hooks for 
 * Left Top, Left Bottom, Right Top and Right Bottom screens.
 * 
 * This also listens to testsuite creation and echoes out for example. 
 * 
 * Class ExcelImportPlugin
 */
class ExcelImportPlugin extends TestlinkPlugin
{
  function _construct()
  {

  }

  function register()
  {
    $this->name = 'ExcelImport';
    $this->description = 'ExcelImport Plugin';

    $this->version = '1.0';

    $this->author = 'iThingsLab';
    $this->contact = 'info@ithings-lab.co.jp';
    $this->url = 'http://www.ithings-lab.co.jp';
  }

  function config()
  {
    return array(
      'config1' => '',
      'config2' => 0
    );
  }

  function hooks()
  {
    $hooks = array(
      'EVENT_RIGHTMENU_BOTTOM' => 'bottom_link',
    );
    return $hooks;
  }

  function bottom_link()
  {
	  $tLink['href'] = 'http://'.$_SERVER['HTTP_HOST'].'/excelutil/';
	  $tLink['label'] = plugin_lang_get('right_bottom_link');
    return $tLink;
  }

}

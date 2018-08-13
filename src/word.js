const async = require ( 'async' )
const officegen = require('officegen')
const fs = require('fs')
const path = require('path')
const os = require('os')
const util = require('./utils')

// var themeXml = fs.readFileSync ( path.resolve ( __dirname, 'themes/testTheme.xml' ), 'utf8' );

module.exports.shengcDocx = function (options) {
  return new Promise(resolve => {
    var defaultOption = {
      contractno: '123456'
    }
  
    var allOptions = Object.assign({}, defaultOption, options)
  
    var docx = officegen ( {
      type: 'docx',
      orientation: 'portrait',
      pageMargins: { top: 10.500, left: 10.500, bottom: 10.500, right: 10.500 }
      // The theme support is NOT working yet...
      // themeXml: themeXml
    } );
    
    // Remove this comment in case of debugging Officegen:
    // officegen.setVerboseMode ( true );
    
    docx.on ( 'error', function ( err ) {
      console.log ( err );
    });
  
    var table = [
      ['合同编号',`${allOptions.contractno}`],
      ['签订地点','江西永修'],
      ['签订日期',`${util.dateFor(allOptions.condate)}`],
    ]
    
    var tableStyle = {
      tableColWidth: 2200,
      tableSize: 24,
      tableColor: "ada",
      tableAlign: "left",
      // align: "center",
    }
    
    console.log(docx)
  
    // 表格
    docx.createTable (table, tableStyle);
  
    // 空行
    docx.createP ({align: 'center'})
    docx.createP ({align: 'center'})
    docx.createP ({align: 'center'})
    docx.createP ({align: 'center'})
    docx.createP ({align: 'center'})
  
    var pObj = docx.createP({align: 'center'})
    pObj.addText('物资采购合同', {font_face: '宋体', font_size: 40, bold: true})
  
    // 9 空行
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
  
    var pObj = docx.createP()
    pObj.addText(`      甲方（购货方）：`, {font_face: '宋体', font_size: 15})
    pObj.addText(`江西中城通达新能源装备有限公司`, {font_face: '宋体', font_size: 15, underline: true})

    docx.createP ({ align: 'center' })
    var pObj = docx.createP()
    pObj.addText(`      乙方（销售方）：`, {font_face: '宋体', font_size: 15})
    pObj.addText(`${allOptions.seller}`, {font_face: '宋体', font_size: 15, underline: true})
    
    // 空行
    docx.createP ({ align: 'center' })
    docx.createP ({ align: 'center' })
  
    var pObj = docx.createP({align: "center"})
    pObj.addText(`物资采购合同`, {font_face: '宋体', font_size: 22, bold: true})
  
    var pObj = docx.createP()
    pObj.addText('购货方（简称甲方）：', {font_face: '宋体', font_size: 10.5, bold: true})
    pObj.addText(`江西中城通达新能源装备有限公司`, {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('销售方（简称乙方）：', {font_face: '宋体', font_size: 10.5, bold: true})
    pObj.addText(`${allOptions.seller}`, {font_face: '宋体', font_size: 10.5})
  
    // 空行
    docx.createP ({ align: 'center' })
  
    var pObj = docx.createP()
    pObj.addText('一、物资清单', {font_face: '宋体', font_size: 10.5})
  
    var wuziTable = [
      ['序号', '产品名称', '规格型号', '金额(元)', '计划来源'],
      ['1', '', '', '', ''],
      ['2', '', '', '', ''],
      ['3', '', '', '', ''],
      ['合计', '', '', '', '']
    ]
  
    var wuziTableStyle = {
      tableColWidth: 1200,
      tableSize: 24,
      tableColor: "ada",
      tableAlign: "center",
      borders: true
    }
  
    // 表格
    docx.createTable (wuziTable, wuziTableStyle)
  
    // 空行
    docx.createP ({ align: 'center' })
  
    var pObj = docx.createP()
    pObj.addText('合同金额：人民币含税金额', {font_size: 10.5 })
    pObj.addText(`${util.formatterNumber(allOptions.connumber)}元`, { font_face: '宋体',bold: true, underline: true, font_size: 10.5})
    pObj.addText('（含', {font_size: 10.5 })
    pObj.addText('16% 增值税', { bold: true, underline: true, font_size: 10.5 })
    pObj.addText('）； 大写：', {font_size: 10.5 })
    pObj.addText(` 人民币${util.convertCurrency(allOptions.connumber)}`, { font_face: '宋体', bold: true, underline: true, font_size: 10.5 })

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    let standId = allOptions.standid
    pObj.addText('二、质量要求技术标准按：', {font_face: '宋体', font_size: 10.5, bold: true})
    pObj.addText('（注：采用的请选择在“□”划“√”，质量标准描述尽量详细， 在横线上填写具体的标准号或附件名）', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('□国家标准', {font_face: '宋体', font_size: 10.5})
    if (standId == "1") {
      pObj.addText('        .', {font_face: '宋体', font_size: 10.5, underline: true, color: "000000"}) 
    }
    var pObj = docx.createP()
    pObj.addText('□行业标准', {font_face: '宋体', font_size: 10.5})
    if (standId == "2") {
      pObj.addText('        .', {font_face: '宋体', font_size: 10.5, underline: true, color: "000000"}) 
    }
    var pObj = docx.createP()
    pObj.addText('□甲方要求标准（见附件）', {font_face: '宋体', font_size: 10.5})

    // 空行
    // docx.createP ({ align: 'center' })
    
    var pObj = docx.createP()
    let balancetype = allOptions.balancetype
    let saletype = allOptions.saletype
    pObj.addText('三、交货地点及运输费用：', {font_face: '宋体', font_size: 10.5, bold: true})
    // 1 先货后款，货已验收合格
    if (balancetype == "1") {
      var pObj = docx.createP()
      pObj.addText('1、交货地点：', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('江西省九江市永修县城南工业园永昌大道39号 。', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      var pObj = docx.createP()
      pObj.addText('2、运输费用：', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('费用由', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('费用由', {font_face: '宋体', font_size: 10.5, bold: true})
      if (saletype == "1") {
        pObj.addText('乙', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      } else {
        pObj.addText('甲', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      }
      pObj.addText('方承担，已包含在合同总金额中。货物需要保险的，保险费用乙方承担。', {font_face: '宋体', font_size: 10.5, bold: true})
    } else if (balancetype == "2") {
      var pObj = docx.createP()
      pObj.addText('1、交货时间：', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText(`${allOptions.paydate.split("-")[0]}年${allOptions.paydate.split("-")[1]}月${allOptions.paydate.split("-")[2]}日前。`, {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      var pObj = docx.createP()
      pObj.addText('2、交货地点：', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('江西省九江市永修县城南工业园永昌大道39号 。', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      var pObj = docx.createP()
      pObj.addText('3、运输费用：', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('费用由', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('费用由', {font_face: '宋体', font_size: 10.5, bold: true})
      if (saletype == "1") {
        pObj.addText('甲', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      } else {
        pObj.addText('乙', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      }
      pObj.addText('方承担，已包含在合同总金额中。货物需要保险的，保险费用乙方承担。', {font_face: '宋体', font_size: 10.5, bold: true})
    } else if (balancetype == "3") {
      var pObj = docx.createP()
      pObj.addText('1、交货时间：', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('${allOptions.paydate.split("-")[0]}年${allOptions.paydate.split("-")[1]}月${allOptions.paydate.split("-")[2]}日前。', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      var pObj = docx.createP()
      pObj.addText('2、交货地点：', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('江西省九江市永修县城南工业园永昌大道39号 。', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      var pObj = docx.createP()
      pObj.addText('3、运输费用：', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('费用由', {font_face: '宋体', font_size: 10.5, bold: true})
      pObj.addText('费用由', {font_face: '宋体', font_size: 10.5, bold: true})
      if (saletype == "1") {
        pObj.addText('甲', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      } else {
        pObj.addText('乙', {font_face: '宋体', font_size: 10.5, bold: true, underline: true})
      }
      pObj.addText('方承担，已包含在合同总金额中。货物需要保险的，保险费用乙方承担。', {font_face: '宋体', font_size: 10.5, bold: true})
    }

    // 空行
    // docx.createP ({ align: 'center' })
    
    var pObj = docx.createP()
    pObj.addText('四、合理损耗及计算方法：', {font_face: '宋体', font_size: 10.5, bold: true})
    var pObj = docx.createP()
    pObj.addText('    以甲方在交货地点验收产品的数量、质量等为准，货物运输损耗由乙方负担，所有合同产品交付甲方后，产品的所有权及产品毁损、灭失的风险转移至甲方。', {font_face: '宋体', font_size: 10.5})

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    pObj.addText('五、产品包装、标识：', {font_face: '宋体', font_size: 10.5, bold: true})
    var pObj = docx.createP()
    pObj.addText('    1、产品的包装、标识由乙方负责,必须采用可以再回收利用或经济处理的环保包装。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    2、包装和标识物由乙方提供，因包装和标识产生的费用由乙方承担，包装和标识必须严格执行相关技术标准，满足产品防尘防锈、防潮防湿、防磕碰损坏等基本要求。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    3、标识清楚、醒目、规范，确保货物储运过程的安全。', {font_face: '宋体', font_size: 10.5})

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    pObj.addText('六、检查及验收：', {font_face: '宋体', font_size: 10.5, bold: true})
    var pObj = docx.createP()
    pObj.addText('    1、产品的检查及验收由甲方按照本合同第二条规定的质量要求、技术标准进行。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    2、乙方必须在交付产品之前或同时向甲方免费提供产品的生产厂家名称、规格型号、基本性能及技术图纸、操作使用与维修保养、出厂检验报告或质量证明书、产品合格证等相关资料。', {font_face: '宋体', font_size: 10.5})

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    pObj.addText('七、质量保证及售后服务：', {font_face: '宋体', font_size: 10.5, bold: true})
    var pObj = docx.createP()
    pObj.addText('    1、质保期：', {font_face: '宋体', font_size: 10.5})
    pObj.addText(`${allOptions.quilty}`, {font_face: '宋体', font_size: 10.5, underline: true})
    pObj.addText('个月', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    2、质保期内甲方在使用过程中如发现产品存在缺陷或技术性能未达到要求，乙方必须在接到甲方通知（口头或书面）后按甲方进度要求及时对产品进行跟踪服务并进行修理、更换或赔偿，每延迟一天，应支付合同总金额0.5%作为违约金。 如果该缺陷出现在质量保证期内，则同时延长所更换物资的质量保证期。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    3、如因乙方产品在使用过程中出现质量问题延误甲方生产进度所造成的相关损失和质量索赔费用由乙方完全负责。 若乙方在收到甲方索赔通知后14个自然日内未予回复，该索赔要求将视为被乙方接受。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    4、因乙方产品质量问题给甲方或第三方造成人身伤害、财产损失的，乙方除应承担相应的赔偿责任和违约责任外，还应承担其他相关的全部损害赔偿和补偿责任。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    5、甲乙双方因不可抗力因素不能履行本合同义务时，均不承担责任，但受不可抗力一方应立即就不可抗力事件对本合同履行的影响通知另一方，同时积极协商应对办法，并有责任采取措施促使合同的履行。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    6、其他情况按《合同法》要求执行。', {font_face: '宋体', font_size: 10.5})

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    // 预付款 1 无 2 有 (默认无)
    let ywyfktype = allOptions.ywyfktype
    pObj.addText('八、结算方式：', {font_face: '宋体', font_size: 10.5, bold: true})

    // 先货后款 已验收
    if (balancetype == "1") {
      var pObj = docx.createP()
      pObj.addText('    1、产品经甲方检验验证合格后，凭增值税票等有效凭证和入库单办理付款流程,并要求在增值税票货物名称及备注栏分别注明物资编码、合同编号。', {font_face: '宋体', font_size: 10.5})
      var pObj = docx.createP()
      pObj.addText('    2、甲方收到乙方开具的税率为16％的增值税专用发票等有效凭证后15个工作日后安排支付货款。乙方违约未按时支付违约金或提供的增值税专用发票审验不合格，甲方有权拒绝支付货款。', {font_face: '宋体', font_size: 10.5})
    } else if (balancetype == "2") {
      if (ywyfktype == "1") {
        var pObj = docx.createP()
        pObj.addText('    1、产品经甲方检验验证合格后，凭增值税票等有效凭证和入库单办理付款流程,并要求在增值税票货物名称及备注栏分别注明物资编码、合同编号。', {font_face: '宋体', font_size: 10.5})
        var pObj = docx.createP()
        pObj.addText('    2、甲方收到乙方开具的税率为16％的增值税专用发票等有效凭证后15个工作日后安排支付货款。乙方违约未按时支付违约金或提供的增值税专用发票审验不合格，甲方有权拒绝支付货款。', {font_face: '宋体', font_size: 10.5})
      } else {
        var pObj = docx.createP()
        pObj.addText('    1、合同签订后甲方支付乙方', {font_face: '宋体', font_size: 10.5})
        pObj.addText('    ', {font_face: '宋体', font_size: 10.5, underline: true})
        pObj.addText(' %预付款。', {font_face: '宋体', font_size: 10.5})
        var pObj = docx.createP()
        pObj.addText('    2、产品经甲方检验验证合格后，凭增值税票等有效凭证和入库单办理付款流程,并要求在增值税票货物名称及备注栏分别注明物资编码、合同编号。', {font_face: '宋体', font_size: 10.5})
        var pObj = docx.createP()
        pObj.addText('    3、甲方收到乙方开具的税率为16％的增值税专用发票等有效凭证后15个工作日后安排支付尾款。乙方违约未按时支付违约金或提供的增值税专用发票审验不合格，甲方有权拒绝支付尾款。', {font_face: '宋体', font_size: 10.5})
      }
    } else if (balancetype == "3") {
      if (ywyfktype == "1") {
        var pObj = docx.createP()
        pObj.addText('    1、产品经甲方检验验证合格后，凭增值税票等有效凭证和入库单办理付款流程,并要求在增值税票货物名称及备注栏分别注明物资编码、合同编号。', {font_face: '宋体', font_size: 10.5})
        var pObj = docx.createP()
        pObj.addText('    2、甲方收到乙方开具的税率为16％的增值税专用发票等有效凭证后15个工作日后安排支付货款。乙方违约未按时支付违约金或提供的增值税专用发票审验不合格，甲方有权拒绝支付货款。', {font_face: '宋体', font_size: 10.5})
      } else {
        var pObj = docx.createP()
        pObj.addText('    1、合同签订后甲方支付乙方', {font_face: '宋体', font_size: 10.5})
        pObj.addText('    ', {font_face: '宋体', font_size: 10.5, underline: true})
        pObj.addText(' %预付款。', {font_face: '宋体', font_size: 10.5})
        var pObj = docx.createP()
        pObj.addText('    2、产品经甲方检验验证合格后，凭增值税票等有效凭证和入库单办理付款流程,并要求在增值税票货物名称及备注栏分别注明物资编码、合同编号。', {font_face: '宋体', font_size: 10.5})
        var pObj = docx.createP()
        pObj.addText('    3、甲方收到乙方开具的税率为16％的增值税专用发票等有效凭证后15个工作日后安排支付尾款。乙方违约未按时支付违约金或提供的增值税专用发票审验不合格，甲方有权拒绝支付尾款。', {font_face: '宋体', font_size: 10.5})
      }
    }

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    pObj.addText('九、知识产权的保护：', {font_face: '宋体', font_size: 10.5, bold: true})
    var pObj = docx.createP()
    pObj.addText('    1、乙方应保证合同产品不会出现任何知识产权瑕疵，如出现第三方对甲方或最终用户提出知识产权索赔和（或）发生行政处罚及其它不利后果，乙方应承担甲方因此遭受的索赔赔偿及相关一切损失，包括但不限于律师费用、诉讼费用等，甲方不承担任何连带责任，所产生的损失全部由乙方承担。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    2、专为甲方设计的产品及其设计,知识产权应属于甲乙双方共同拥有，未经甲方书面同意，乙方不得申请专利及不得将该知识产权及产品、设计卖与第三方。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    3、甲乙双方终止合同关系，乙方应将技术图纸资料及其他记录相关信息的载体交给甲方，乙方不得以复制件，电子文档等任何形式保留备份。未经甲方书面同意不得向第三方提供相关资料，不得对图纸做任何修改。', {font_face: '宋体', font_size: 10.5})


    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    pObj.addText('十、违约责任：', {font_face: '宋体', font_size: 10.5, bold: true})
    var pObj = docx.createP()
    pObj.addText('    1、乙方提供的产品品种、规格型号不符、包装不当或其他质量问题，由乙方按甲方要求负责处理，所发生的费用由乙方承担。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    2、乙方延期交货时，按照合同交货进度要求，每延期一天按本合同总金额的0.5%支付违约金，并承担甲方因此所受的损失费用。如交货延期达到30日，甲方有权解除合同，由此产生的一切损失由乙方承担。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    3、乙方提前多交的或不合格的产品由甲方代管，乙方应向甲方支付代管期内的保管、保养等费用并赔偿甲方因此发生的损失。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    4、其他情况按《合同法》的有关规定要求执行。', {font_face: '宋体', font_size: 10.5})

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    pObj.addText('十一、解决合同纠纷方式：', {font_face: '宋体', font_size: 10.5, bold: true})
    pObj.addText('双方协商解决，协商不成，向甲方所在地法院提起诉讼。', {font_face: '宋体', font_size: 10.5})

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    pObj.addText('十二、其他约定事项：', {font_face: '宋体', font_size: 10.5, bold: true})
    var pObj = docx.createP()
    pObj.addText('    1、双方特别申明并一致同意，在本合同履行期内若市场上同类产品或可替代产品的降价幅度为单项产品价格5%以上（含5%），买方有权选择单方面解除本合同或中止履行本合同，并选择以该市场价格直接向其他生产商或销售商采购上述产品，也可以与乙方以该市场价格重新签定合同。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    2、本合同有关的质量保证协议、预留质保金协议等协议与本合同具有同等法律效力,本合同未涉及事宜按有关法律办理。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    3、合同的变更、中途终止要双方书面形式确认才生效。', {font_face: '宋体', font_size: 10.5})
    var pObj = docx.createP()
    pObj.addText('    4、乙方车辆人员进入甲方施工作业现场，必须遵守甲方现场相关的安全制度，服从现场管理人员的指挥，必须自行配备安全帽，否则，在现场发生的人员伤害和财产损失由乙方承担。乙方钢材在场内、外运输及卸车过程中若造成甲方或第三者人员伤亡或财产损失，均由乙方承担相关责任。', {font_face: '宋体', font_size: 10.5})

    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP()
    pObj.addText('十三、合同有效期：', {font_face: '宋体', font_size: 10.5, bold: true})
    var pObj = docx.createP()
    pObj.addText('    本合同自双方签字盖章之日起生效。本合同正本一式肆份，甲方叁份，乙方壹份。', {font_face: '宋体', font_size: 10.5})
    
    // 空行
    // docx.createP ({ align: 'center' })
    // 空行
    // docx.createP ({ align: 'center' })
    // 空行
    // docx.createP ({ align: 'center' })
    // 空行
    // docx.createP ({ align: 'center' })
    // 空行
    // docx.createP ({ align: 'center' })

    var pObj = docx.createP({align: 'center'})
    pObj.addText('（以下无正文）', {font_face: '宋体', font_size: 10.5})

    // 空行
    docx.createP ({ align: 'center' })
    // 空行
    docx.createP ({ align: 'center' })
    // 空行
    docx.createP ({ align: 'center' })
    // 空行
    docx.createP ({ align: 'center' })

    var neirongTable = [
      ['甲方（盖章）：江西中城通达新能源装备有限公司', '乙方：'],
      ['地址：江西省九江市永修县城南工业园39号', '地址：'],
      ['邮政编码：330300', '邮政编码：'],
      ['电话号码：', '地址：'],
      ['地址：江西省九江市永修县城南工业园39号', '电话号码：'],
      ['开户行：中国工商银行永修支行', '开户行：'],
      ['银行帐号：1507255009200093088', '银行帐号：'],
      ['税号：91360425MA36UJ8E9C', '税号：'],
      ['甲方法定代表人：丁刚', '乙方法定代表人：'],
      ['甲方经办人:', '乙方经办人:']
    ]
  
    var neirongTableStyle = {
      tableColWidth: 4600,
      tableSize: 10,
      tableColor: "ada",
      tableAlign: "left"
    }

    // 表格
    docx.createTable (neirongTable, neirongTableStyle)
  
    // 文档输出
    var out = fs.createWriteStream (path.join(os.tmpdir(), 'out.docx'))
    
    out.on ( 'error', function ( err ) {
      console.log ( err );
    });
    
    async.parallel ([
      function ( done ) {
        out.on ( 'close', function () {
          console.log ( 'Finish to create a DOCX file.' );
          done ( null );
        });
        docx.generate ( out )
        resolve()
      }
    
    ], function ( err ) {
      if ( err ) {
        console.log ( 'error: ' + err );
      } // Endif.
    });
  })
}

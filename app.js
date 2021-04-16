const fs = require('fs')
const path = require('path')
const XLSX = require('xlsx')
const moment = require('moment')

// 转XLSX
const toXLSX = (_data, name) => {
	try {
		const _headers = [
			{ key: 'date_time', title: '时间' },
			{ key: 'nick_name', title: '昵称' },
			{ key: 'content', title: '问题' },
			{ key: 'has_reply', title: '是否回复' },
			// { key: 'fakeid', title: '用户id' },
			// { key: 'remark', title: '备注' }
		]
		const headers = _headers.map(({ title }) => title)
			// 为 _headers 添加对应的单元格位置
			// [ { v: 'id', position: 'A1' },
			//   { v: 'name', position: 'B1' },
			//   { v: 'age', position: 'C1' },
			//   { v: 'country', position: 'D1' },
			//   { v: 'remark', position: 'E1' } ]
			.map((v, i) => Object.assign({}, { v: v, position: String.fromCharCode(65 + i) + 1 }))
			// 转换成 worksheet 需要的结构
			// { A1: { v: 'id' }, 
			//   B1: { v: 'name' },
			//   C1: { v: 'age' },
			//   D1: { v: 'country' },
			//   E1: { v: 'remark' } }
			.reduce((prev, next) => Object.assign({}, prev, {
				[next.position]: { v: next.v }
			}), {})

		const data = _data
			// 匹配 headers 的位置，生成对应的单元格数据
			// [ [ { v: '1', position: 'A2' },
			//     { v: 'test1', position: 'B2' },
			//     { v: '30', position: 'C2' },
			//     { v: 'China', position: 'D2' },
			//     { v: 'hello', position: 'E2' } ],
			//   [ { v: '2', position: 'A3' },
			//     { v: 'test2', position: 'B3' },
			//     { v: '20', position: 'C3' },
			//     { v: 'America', position: 'D3' },
			//     { v: 'world', position: 'E3' } ],
			//   [ { v: '3', position: 'A4' },
			//     { v: 'test3', position: 'B4' },
			//     { v: '18', position: 'C4' },
			//     { v: 'Unkonw', position: 'D4' },
			//     { v: '???', position: 'E4' } ] ]
			.map((v, i) => _headers.map(({ key }, j) => Object.assign({}, { v: v[key], position: String.fromCharCode(65 + j) + (i + 2) })))
			// 对刚才的结果进行降维处理（二维数组变成一维数组）
			// [ { v: '1', position: 'A2' },
			//   { v: 'test1', position: 'B2' },
			//   { v: '30', position: 'C2' },
			//   { v: 'China', position: 'D2' },
			//   { v: 'hello', position: 'E2' },
			//   { v: '2', position: 'A3' },
			//   { v: 'test2', position: 'B3' },
			//   { v: '20', position: 'C3' },
			//   { v: 'America', position: 'D3' },
			//   { v: 'world', position: 'E3' },
			//   { v: '3', position: 'A4' },
			//   { v: 'test3', position: 'B4' },
			//   { v: '18', position: 'C4' },
			//   { v: 'Unkonw', position: 'D4' },
			//   { v: '???', position: 'E4' } ]
			.reduce((prev, next) => prev.concat(next))
			// 转换成 worksheet 需要的结构
			//   { A2: { v: '1' },
			//     B2: { v: 'test1' },
			//     C2: { v: '30' },
			//     D2: { v: 'China' },
			//     E2: { v: 'hello' },
			//     A3: { v: '2' },
			//     B3: { v: 'test2' },
			//     C3: { v: '20' },
			//     D3: { v: 'America' },
			//     E3: { v: 'world' },
			//     A4: { v: '3' },
			//     B4: { v: 'test3' },
			//     C4: { v: '18' },
			//     D4: { v: 'Unkonw' },
			//     E4: { v: '???' } }
			.reduce((prev, next) => Object.assign({}, prev, {
				[next.position]: { v: next.v }
			}), {})

		// 合并 headers 和 data
		const output = Object.assign({}, headers, data)

		// 获取所有单元格的位置
		const outputPos = Object.keys(output)

		// 计算出范围
		const ref = outputPos[0] + ':' + outputPos[outputPos.length - 1]

		// 构建 workbook 对象
		const wb = {
			SheetNames: ['sheet1'],
			Sheets: {
				'sheet1': Object.assign({}, output, { '!ref': ref })
			}
		}
		// 导出 Excel
		XLSX.writeFile(wb, `${ name }.xlsx`)
		// return result.data
	} catch (e) {
		console.log('--------->xlsx处理错误', e)
	}
}

// 解析目录
let analysisDirectory = async () => {
	let path = './txts/'
	let list = []
	await new Promise((reslove, reject) => {
		let fs = require('fs')
		// 解析文件夹
		// console.log('--------->解析文件夹', path)
		fs.readdir(path, (err, res) => {
			if (err) {
				console.log('--------->err', err)
				return this.error('文件读取错误', e, -2003)
			}
			// console.log('--------->读取到当前文件夹', res)
			// 解析文件夹的每一个文件
			for (let file of res) {
				console.log('--------->解析文件', `${ path }/${ file }`)
				let fileData = fs.readFileSync(`${ path }/${ file }`)
				let str = fileData.toString()
				let strArr = str.split('{"base_resp":')
				// console.log('--------->str', str)
				for (let item of strArr) {
					if (!item) continue
					let strObj = JSON.parse(`{"base_resp":${ item }`)
					list = list.concat(strObj.item)
				}
				// console.log('--------->_list', list)
				// console.log('--------->_list', list.length)
			}
			reslove()
		})
	})
	console.log(`--------->正在导出${ list.length }条`, )
	let arrList = []
	// 处理数据
	for (let i of list) {
		let { biz_last_replay_id, can_replay, is_blacked, msg_items } = i.msg
		msg_items = JSON.parse(msg_items)
		// 把所有数据整合成一个对象
		for (let msg of msg_items.msg_item) {
			let timeStamp = msg.date_time // 保留时间戳后续排序
			let time = msg.date_time * 1000
			msg.date_time = moment(time).format('YYYY-MM-DD HH:mm')
			msg.has_reply = msg.has_reply === 0 ? '' : '已回复'
			msg.content = msg.content ? msg.content : '(非文本内容)'
			msg = {
				biz_last_replay_id,
				can_replay,
				is_blacked,
				timeStamp,
				...msg
			}
			arrList.push(msg)
		}
	}
	arrList = arrList.sort((a, b) => b.timeStamp - a.timeStamp) // 处理整个数组进行时间排序
	const _data = [{
		id: '1',
		'nick_name': 'test1',
		content: '30',
		country: 'China',
	}, {
		id: '2',
		nick_name: 'test2',
		content: '20',
		country: 'America',
		remark2: 'world'
	}, {
		id: '3',
		nick_name: 'test3',
		content: '18',
		country: 'Unkonw',
		remark: '???'
	}]
	let date = moment().format('YYYY-MM-DD')
	let name = `成都供电公众号会话导出${ date }`
	toXLSX(arrList, name) //使用业务逻辑层的方法返回值
	console.log('--------->导出完成', )
}
analysisDirectory(null, true) // 执行
/**
 * 			佛曰:
 * 				写字楼里写字间，写字间里程序员；
 * 				程序人员写程序，又拿程序换酒钱。
 * 				酒醒只在网上坐，酒醉还来网下眠；
 * 				酒醉酒醒日复日，网上网下年复年。
 * 				但愿老死电脑间，不愿鞠躬老板前；
 * 				奔驰宝马贵者趣，公交自行程序员。
 * 				别人笑我忒疯癫，我笑自己命太贱；
 * 				不见满街漂亮妹，哪个归得程序员？
 *
 * @description 批量剪切图片 把图片放入config.filePath(可随意存放任意文件,自动过滤无效文件,并保存至config.surplusImage),执行: node.app 
 * @tutorial 暂无参考文档
 * @param {String} paramsName = 未知的参数
 * @event 暂无事件
 * @example 暂无示例
 * @return {String} {暂无返回值}
 * @author Breathe
 */

import express from 'express';
import exceljs from 'exceljs';
import path from 'path';
import axios from 'axios';
import 'dotenv/config';

const app = express();

interface ResponseData {
  anoCalendario: number;
  precatorio: string;
  nomeCredor: string;
  nomeBeneficiarioPrevidencia: string;
  nomeBeneficiarioImpostoRenda: string;
  nomeBeneficiarioISSQN: string;
  cpfCnpjCredor: string;
  cpfCnpjBeneficiarioPrevidencia: string;
  cpfCnpjBeneficiarioImpostoRenda: string;
  cpfCnpjBeneficiarioISSQN: string;
  valorBruto: number;
  valorImpostoRenda: number;
  valorPrevidencia: number;
  valorISSQN: number;
  valorLiquido: number;
  mesLiquidacao: number;
  dataDoPagamento: string;
  mesesRRA: number;
};

app.get('/', async (request, response) => {
  const workbook = new exceljs.Workbook();
  const worksheet = workbook.addWorksheet('Data');

  const { data } = await axios.get<ResponseData[]>(process.env.APP_API_URL || '', {
    auth: {
      username: process.env.API_USER || '',
      password: process.env.API_PASSWORD || '',
    }
  });

  worksheet.columns = [
    { header: 'Ano Calendário', key: 'anoCalendario' },
    { header: 'Precatório', key: 'precatorio' },
    { header: 'Nome Credor', key: 'nomeCredor' },
    { header: 'Nome Beneficiário Previdência', key: 'nomeBeneficiarioPrevidencia' },
    { header: 'Nome Beneficiário IR', key: 'nomeBeneficiarioImpostoRenda' },
    { header: 'Nome Beneficiário ISSQN', key: 'nomeBeneficiarioISSQN'},
    { header: 'CPF/CNPJ Credor', key: 'cpfCnpjCredor' },
    { header: 'CPF/CNPJ Beneficiário Previdência', key: 'cpfCnpjBeneficiarioPrevidencia' },
    { header: 'CPF/CNPJ Beneficário IR', key: 'cpfCnpjBeneficiarioImpostoRenda' },
    { header: 'CPF/CNPJ Beneficiário ISSQN', key: 'cpfCnpjBeneficiarioISSQN' },
    { header: 'Valor Bruto', key: 'valorBruto' },
    { header: 'Valor IR', key: 'valorImpostoRenda' },
    { header: 'Valor Previdência', key: 'valorPrevidencia' },
    { header: 'Valor ISSQN', key: 'valorISSQN' },
    { header: 'Valor Líquido', key: 'valorLiquido' },
    { header: 'Mês Liquidação', key: 'mesLiquidacao' },
    { header: 'Data Pagamento', key: 'dataDoPagamento' },
    { header: 'Meses RRA', key: 'mesesRRA' },
  ];

  data.map(item => {
    worksheet.addRow({
      anoCalendario: item.anoCalendario,
      precatorio: item.precatorio,
      nomeCredor: item.nomeCredor,
      nomeBeneficiarioPrevidencia: item.nomeBeneficiarioPrevidencia,
      nomeBeneficiarioImpostoRenda: item.nomeBeneficiarioImpostoRenda,
      nomeBeneficiarioISSQN: item.nomeBeneficiarioISSQN,
      cpfCnpjCredor: item.cpfCnpjCredor,
      cpfCnpjBeneficiarioPrevidencia: item.cpfCnpjBeneficiarioPrevidencia,
      cpfCnpjBeneficiarioImpostoRenda: item.cpfCnpjBeneficiarioImpostoRenda,
      cpfCnpjBeneficiarioISSQN: item.cpfCnpjBeneficiarioISSQN,
      valorBruto: item.valorBruto,
      valorImpostoRenda: item.valorImpostoRenda,
      valorPrevidencia: item.valorPrevidencia,
      valorISSQN: item.valorISSQN,
      valorLiquido: item.valorLiquido,
      mesLiquidacao: item.mesLiquidacao,
      dataDoPagamento: item.dataDoPagamento,
      mesesRRA: item.mesesRRA,
    });
  });

  const fileName = `${new Date().getTime()}.xlsx`;

  await workbook.xlsx.writeFile(fileName);

  response.setHeader('Content-disposition', `attachment; filename=${fileName}`);
  response.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  return response.download(path.join(__dirname, '..', fileName));
})

app.listen('3333', () => console.log('Server started on port 3333'));
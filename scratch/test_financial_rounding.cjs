function roundMoney(val) {
  return Math.round((Number(val || 0) + 0.00001) * 100) / 100;
}

const rep1v = 1582.30;
const lub1v = 1427.45;
const mo1v  = 2609.40;
const prev1v = rep1v + lub1v + mo1v;

const rep3v = roundMoney(rep1v * 3);
const lub3v = roundMoney(lub1v * 3);
const mo3v  = roundMoney(mo1v * 3);
const prev3v = roundMoney(prev1v * 3);
const corr3v = roundMoney(prev3v * 0.30);
const granTot = roundMoney(prev3v + corr3v);

console.log('Total Repuestos 3V:   ', rep3v);
console.log('Total Lubricantes 3V: ', lub3v);
console.log('Total Mano Obra 3V:   ', mo3v);
console.log('Total Preventivo 3V:  ', prev3v);
console.log('Mantenimiento Correctivo (30%):', corr3v);
console.log('Gran Total (1+2+3+4): ', granTot);

console.log('\nMatches Target Exactly:');
console.log('Preventivo: $16,857.45 =>', prev3v === 16857.45);
console.log('Correctivo: $5,057.24  =>', corr3v === 5057.24);
console.log('Gran Total: $21,914.69 =>', granTot === 21914.69);

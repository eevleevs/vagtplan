import { assertEquals } from '@std/assert'

import { main } from '../vagtplan.ts'

for (const { name } of Deno.readDirSync('tests')) {
  if (name.startsWith('~$') || !name.endsWith('.xlsx')) continue
  Deno.test(name, () =>
    assertEquals(
      main(`tests/${name}`),
      JSON.parse(
        Deno.readTextFileSync(`tests/${name}`.replace(/.xlsx$/, '.json')),
      ),
    ))
}

import { assertEquals } from '@std/assert'

import { main } from '../vagtplan.ts'

Deno.test('VAGTSKEMA MARTS 2025 (002)', () =>
  assertEquals(
    main('tests/VAGTSKEMA MARTS 2025 (002).xlsx'),
    JSON.parse(Deno.readTextFileSync('tests/VAGTSKEMA MARTS 2025 (002).json')),
  ))

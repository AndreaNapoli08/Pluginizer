// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna
import * as React from 'react';
import { TextBlocks } from './Blocks/TextBlocks'
import { BulletList } from './Blocks/BulletList'
import { OrderedList} from './Blocks/OrderedList'
import { Nesting } from './Blocks/Nesting'

export const Blocks = () => {
    return (
        <div>
            <TextBlocks />
            <hr />
            <BulletList />
            <hr />
            <OrderedList />
            <hr />
            <Nesting />
        </div>
    )
}
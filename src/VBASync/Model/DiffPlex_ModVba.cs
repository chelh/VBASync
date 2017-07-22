/* This file contains various parts of the DiffPlex package, modified for use with VBA.
 * All modifications to the DiffPlex code are released into the public domain.
 *
 * DiffPlex is distributed under the Apache License, included below:
 *
 * * Apache License
 * * Version 2.0, January 2004
 * * http://www.apache.org/licenses/
 * *
 * * TERMS AND CONDITIONS FOR USE, REPRODUCTION, AND DISTRIBUTION
 * *
 * * 1. Definitions.
 * *
 * * "License" shall mean the terms and conditions for use, reproduction, and distribution as defined by Sections 1 through 9 of this document.
 * *
 * * "Licensor" shall mean the copyright owner or entity authorized by the copyright owner that is granting the License.
 * *
 * * "Legal Entity" shall mean the union of the acting entity and all other entities that control, are controlled by, or are under common control
 * * with that entity. For the purposes of this definition, "control" means (i) the power, direct or indirect, to cause the direction or management
 * * of such entity, whether by contract or otherwise, or (ii) ownership of fifty percent (50%) or more of the outstanding shares, or
 * * (iii) beneficial ownership of such entity.
 * *
 * * "You" (or "Your") shall mean an individual or Legal Entity exercising permissions granted by this License.
 * *
 * * "Source" form shall mean the preferred form for making modifications, including but not limited to software source code, documentation source,
 * * and configuration files.
 * *
 * * "Object" form shall mean any form resulting from mechanical transformation or translation of a Source form, including but not limited to
 * * compiled object code, generated documentation, and conversions to other media types.
 * *
 * * "Work" shall mean the work of authorship, whether in Source or Object form, made available under the License, as indicated by a copyright notice
 * * that is included in or attached to the work (an example is provided in the Appendix below).
 * *
 * * "Derivative Works" shall mean any work, whether in Source or Object form, that is based on (or derived from) the Work and for which the editorial
 * * revisions, annotations, elaborations, or other modifications represent, as a whole, an original work of authorship. For the purposes of this License,
 * * Derivative Works shall not include works that remain separable from, or merely link (or bind by name) to the interfaces of, the Work and
 * * Derivative Works thereof.
 * *
 * * "Contribution" shall mean any work of authorship, including the original version of the Work and any modifications or additions to that Work or
 * * Derivative Works thereof, that is intentionally submitted to Licensor for inclusion in the Work by the copyright owner or by an individual or
 * * Legal Entity authorized to submit on behalf of the copyright owner. For the purposes of this definition, "submitted" means any form of electronic,
 * * verbal, or written communication sent to the Licensor or its representatives, including but not limited to communication on electronic mailing lists,
 * * source code control systems, and issue tracking systems that are managed by, or on behalf of, the Licensor for the purpose of discussing and
 * * improving the Work, but excluding communication that is conspicuously marked or otherwise designated in writing by the copyright owner as
 * * "Not a Contribution."
 * *
 * * "Contributor" shall mean Licensor and any individual or Legal Entity on behalf of whom a Contribution has been received by Licensor and subsequently
 * * incorporated within the Work.
 * *
 * * 2. Grant of Copyright License.
 * *
 * * Subject to the terms and conditions of this License, each Contributor hereby grants to You a perpetual, worldwide, non-exclusive, no-charge, royalty-free,
 * * irrevocable copyright license to reproduce, prepare Derivative Works of, publicly display, publicly perform, sublicense, and distribute the Work and such
 * * Derivative Works in Source or Object form.
 * *
 * * 3. Grant of Patent License.
 * *
 * * Subject to the terms and conditions of this License, each Contributor hereby grants to You a perpetual, worldwide, non-exclusive, no-charge, royalty-free,
 * * irrevocable (except as stated in this section) patent license to make, have made, use, offer to sell, sell, import, and otherwise transfer the Work,
 * * where such license applies only to those patent claims licensable by such Contributor that are necessarily infringed by their Contribution(s) alone or
 * * by combination of their Contribution(s) with the Work to which such Contribution(s) was submitted. If You institute patent litigation against any entity
 * * (including a cross-claim or counterclaim in a lawsuit) alleging that the Work or a Contribution incorporated within the Work constitutes direct or
 * * contributory patent infringement, then any patent licenses granted to You under this License for that Work shall terminate as of the date such litigation
 * * is filed.
 * *
 * * 4. Redistribution.
 * *
 * * You may reproduce and distribute copies of the Work or Derivative Works thereof in any medium, with or without modifications, and in Source or Object form,
 * * provided that You meet the following conditions:
 * *
 * * You must give any other recipients of the Work or Derivative Works a copy of this License; and
 * * You must cause any modified files to carry prominent notices stating that You changed the files; and
 * * You must retain, in the Source form of any Derivative Works that You distribute, all copyright, patent, trademark, and attribution notices from the Source
 * *     form of the Work, excluding those notices that do not pertain to any part of the Derivative Works; and
 * * If the Work includes a "NOTICE" text file as part of its distribution, then any Derivative Works that You distribute must include a readable copy of the
 * *     attribution notices contained within such NOTICE file, excluding those notices that do not pertain to any part of the Derivative Works, in at least one of
 * *     the following places: within a NOTICE text file distributed as part of the Derivative Works; within the Source form or documentation, if provided along
 * *     with the Derivative Works; or, within a display generated by the Derivative Works, if and wherever such third-party notices normally appear. The contents
 * *     of the NOTICE file are for informational purposes only and do not modify the License. You may add Your own attribution notices within Derivative Works
 * *     that You distribute, alongside or as an addendum to the NOTICE text from the Work, provided that such additional attribution notices cannot be construed
 * *     as modifying the License.
 * * You may add Your own copyright statement to Your modifications and may provide additional or different license terms and conditions for use, reproduction,
 * * or distribution of Your modifications, or for any such Derivative Works as a whole, provided Your use, reproduction, and distribution of the Work otherwise
 * * complies with the conditions stated in this License.
 * *
 * * 5. Submission of Contributions.
 * *
 * * Unless You explicitly state otherwise, any Contribution intentionally submitted for inclusion in the Work by You to the Licensor shall be under the terms
 * * and conditions of this License, without any additional terms or conditions. Notwithstanding the above, nothing herein shall supersede or modify the terms
 * * of any separate license agreement you may have executed with Licensor regarding such Contributions.
 * *
 * * 6. Trademarks.
 * *
 * * This License does not grant permission to use the trade names, trademarks, service marks, or product names of the Licensor, except as required for
 * * reasonable and customary use in describing the origin of the Work and reproducing the content of the NOTICE file.
 * *
 * * 7. Disclaimer of Warranty.
 * *
 * * Unless required by applicable law or agreed to in writing, Licensor provides the Work (and each Contributor provides its Contributions) on an "AS IS" BASIS,
 * * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied, including, without limitation, any warranties or conditions of TITLE, NON-INFRINGEMENT,
 * * MERCHANTABILITY, or FITNESS FOR A PARTICULAR PURPOSE. You are solely responsible for determining the appropriateness of using or redistributing the Work and
 * * assume any risks associated with Your exercise of permissions under this License.
 * *
 * * 8. Limitation of Liability.
 * *
 * * In no event and under no legal theory, whether in tort (including negligence), contract, or otherwise, unless required by applicable law (such as deliberate
 * * and grossly negligent acts) or agreed to in writing, shall any Contributor be liable to You for damages, including any direct, indirect, special, incidental,
 * * or consequential damages of any character arising as a result of this License or out of the use or inability to use the Work (including but not limited to
 * * damages for loss of goodwill, work stoppage, computer failure or malfunction, or any and all other commercial damages or losses), even if such Contributor
 * * has been advised of the possibility of such damages.
 * *
 * * 9. Accepting Warranty or Additional Liability.
 * *
 * * While redistributing the Work or Derivative Works thereof, You may choose to offer, and charge a fee for, acceptance of support, warranty, indemnity, or other
 * * liability obligations and/or rights consistent with this License. However, in accepting such obligations, You may act only on Your own behalf and on Your sole
 * * responsibility, not on behalf of any other Contributor, and only if You agree to indemnify, defend, and hold each Contributor harmless for any liability
 * * incurred by, or claims asserted against, such Contributor by reason of your accepting any such warranty or additional liability.
 * *
 * * END OF TERMS AND CONDITIONS
 */

using System;
using System.Collections.Generic;
using System.Text;

namespace VBASync.Model
{
    internal enum Edit
    {
        None = 0,
        DeleteRight = 1,
        DeleteLeft = 2,
        InsertDown = 3,
        InsertUp = 4
    }

    internal enum LineState
    {
        New = 0,
        Normal = 1,
        Quoted = 2,
        Comment = 3,
        CommentAwaitE = 4,
        CommentAwaitM = 5,
        CommentAwaitSpace = 6
    }

    internal static class VbaDiffer
    {
        private static readonly string[] emptyStringArray = new string[0];

        public static DiffResult CreateCustomDiffs(string oldText, string newText, bool ignoreWhiteSpace, Func<string, string[]> chunker)
        {
            return CreateCustomDiffs(oldText, newText, ignoreWhiteSpace, false, chunker);
        }

        public static DiffResult CreateCustomDiffs(string oldText, string newText, bool ignoreWhiteSpace, bool ignoreCase, Func<string, string[]> chunker)
        {
            if (oldText == null) throw new ArgumentNullException(nameof(oldText));
            if (newText == null) throw new ArgumentNullException(nameof(newText));
            if (chunker == null) throw new ArgumentNullException(nameof(chunker));

            var pieceHash = new Dictionary<string, int>();
            var lineDiffs = new List<DiffBlock>();

            var modOld = new ModificationData(oldText);
            var modNew = new ModificationData(newText);

            BuildPieceHashes(pieceHash, modOld, ignoreWhiteSpace, ignoreCase, chunker);
            BuildPieceHashes(pieceHash, modNew, ignoreWhiteSpace, ignoreCase, chunker);

            BuildModificationData(modOld, modNew);

            var piecesALength = modOld.HashedPieces.Length;
            var piecesBLength = modNew.HashedPieces.Length;
            var posA = 0;
            var posB = 0;

            do
            {
                while (posA < piecesALength
                       && posB < piecesBLength
                       && !modOld.Modifications[posA]
                       && !modNew.Modifications[posB])
                {
                    posA++;
                    posB++;
                }

                var beginA = posA;
                var beginB = posB;
                for (; posA < piecesALength && modOld.Modifications[posA]; posA++) ;

                for (; posB < piecesBLength && modNew.Modifications[posB]; posB++) ;

                var deleteCount = posA - beginA;
                var insertCount = posB - beginB;
                if (deleteCount > 0 || insertCount > 0)
                {
                    lineDiffs.Add(new DiffBlock(beginA, deleteCount, beginB, insertCount));
                }
            } while (posA < piecesALength && posB < piecesBLength);

            return new DiffResult(modOld.Pieces, modNew.Pieces, lineDiffs);
        }

        public static DiffResult CreateVbaDiffs(string oldText, string newText)
        {
            return CreateCustomDiffs(oldText, newText, false, s => s.Split(new[] { "\r\n" }, StringSplitOptions.None));
        }

        private static void BuildModificationData(ModificationData A, ModificationData B)
        {
            var N = A.HashedPieces.Length;
            var M = B.HashedPieces.Length;
            var MAX = M + N + 1;
            var forwardDiagonal = new int[MAX + 1];
            var reverseDiagonal = new int[MAX + 1];
            BuildModificationData(A, 0, N, B, 0, M, forwardDiagonal, reverseDiagonal);
        }

        private static void BuildModificationData
            (ModificationData A,
             int startA,
             int endA,
             ModificationData B,
             int startB,
             int endB,
             int[] forwardDiagonal,
             int[] reverseDiagonal)
        {
            while (startA < endA && startB < endB && A.HashedPieces[startA].Equals(B.HashedPieces[startB]))
            {
                startA++;
                startB++;
            }
            while (startA < endA && startB < endB && A.HashedPieces[endA - 1].Equals(B.HashedPieces[endB - 1]))
            {
                endA--;
                endB--;
            }

            var aLength = endA - startA;
            var bLength = endB - startB;
            if (aLength > 0 && bLength > 0)
            {
                var res = CalculateEditLength(A.HashedPieces, startA, endA, B.HashedPieces, startB, endB, forwardDiagonal, reverseDiagonal);
                if (res.EditLength <= 0) return;

                if (res.LastEdit == Edit.DeleteRight && res.StartX - 1 > startA)
                    A.Modifications[--res.StartX] = true;
                else if (res.LastEdit == Edit.InsertDown && res.StartY - 1 > startB)
                    B.Modifications[--res.StartY] = true;
                else if (res.LastEdit == Edit.DeleteLeft && res.EndX < endA)
                    A.Modifications[res.EndX++] = true;
                else if (res.LastEdit == Edit.InsertUp && res.EndY < endB)
                    B.Modifications[res.EndY++] = true;

                BuildModificationData(A, startA, res.StartX, B, startB, res.StartY, forwardDiagonal, reverseDiagonal);

                BuildModificationData(A, res.EndX, endA, B, res.EndY, endB, forwardDiagonal, reverseDiagonal);
            }
            else if (aLength > 0)
            {
                for (var i = startA; i < endA; i++)
                    A.Modifications[i] = true;
            }
            else if (bLength > 0)
            {
                for (var i = startB; i < endB; i++)
                    B.Modifications[i] = true;
            }
        }

        private static void BuildPieceHashes(IDictionary<string, int> pieceHash, ModificationData data, bool ignoreWhitespace, bool ignoreCase, Func<string, string[]> chunker)
        {
            var pieces = string.IsNullOrEmpty(data.RawData) ? emptyStringArray : chunker(data.RawData);
            data.Pieces = pieces;
            data.HashedPieces = new int[pieces.Length];
            data.Modifications = new bool[pieces.Length];

            var startInComment = false;
            for (var i = 0; i < pieces.Length; i++)
            {
                var piece = UppercaseVbaSymbols(pieces[i], ref startInComment);
                if (ignoreWhitespace) piece = piece.Trim();
                if (ignoreCase) piece = piece.ToUpperInvariant();

                if (pieceHash.ContainsKey(piece))
                {
                    data.HashedPieces[i] = pieceHash[piece];
                }
                else
                {
                    data.HashedPieces[i] = pieceHash.Count;
                    pieceHash[piece] = pieceHash.Count;
                }
            }
        }

        /// <summary>
        /// Finds the middle snake and the minimum length of the edit script comparing string A and B
        /// </summary>
        /// <param name="A"></param>
        /// <param name="startA">Lower bound inclusive</param>
        /// <param name="endA">Upper bound exclusive</param>
        /// <param name="B"></param>
        /// <param name="startB">lower bound inclusive</param>
        /// <param name="endB">upper bound exclusive</param>
        /// <returns></returns>
        private static EditLengthResult CalculateEditLength(int[] A, int startA, int endA, int[] B, int startB, int endB)
        {
            var N = endA - startA;
            var M = endB - startB;
            var MAX = M + N + 1;

            var forwardDiagonal = new int[MAX + 1];
            var reverseDiagonal = new int[MAX + 1];
            return CalculateEditLength(A, startA, endA, B, startB, endB, forwardDiagonal, reverseDiagonal);
        }

        private static EditLengthResult CalculateEditLength(int[] A, int startA, int endA, int[] B, int startB, int endB, int[] forwardDiagonal, int[] reverseDiagonal)
        {
            if (A == null) throw new ArgumentNullException(nameof(A));
            if (B == null) throw new ArgumentNullException(nameof(B));

            if (A.Length == 0 && B.Length == 0)
            {
                return new EditLengthResult();
            }

            var N = endA - startA;
            var M = endB - startB;
            var MAX = M + N + 1;
            var HALF = MAX / 2;
            var delta = N - M;
            var deltaEven = delta % 2 == 0;
            forwardDiagonal[1 + HALF] = 0;
            reverseDiagonal[1 + HALF] = N + 1;

            for (var D = 0; D <= HALF; D++)
            {
                // forward D-path
                Edit lastEdit;
                for (var k = -D; k <= D; k += 2)
                {
                    var kIndex = k + HALF;
                    int x, y;
                    if (k == -D || (k != D && forwardDiagonal[kIndex - 1] < forwardDiagonal[kIndex + 1]))
                    {
                        x = forwardDiagonal[kIndex + 1]; // y up    move down from previous diagonal
                        lastEdit = Edit.InsertDown;
                    }
                    else
                    {
                        x = forwardDiagonal[kIndex - 1] + 1; // x up     move right from previous diagonal
                        lastEdit = Edit.DeleteRight;
                    }
                    y = x - k;
                    var startX = x;
                    var startY = y;
                    while (x < N && y < M && A[x + startA] == B[y + startB])
                    {
                        ++x;
                        ++y;
                    }

                    forwardDiagonal[kIndex] = x;

                    if (!deltaEven && k - delta >= -D + 1 && k - delta <= D - 1)
                    {
                        var revKIndex = (k - delta) + HALF;
                        var revX = reverseDiagonal[revKIndex];
                        var revY = revX - k;
                        if (revX <= x && revY <= y)
                        {
                            return new EditLengthResult
                            {
                                EditLength = (2 * D) - 1,
                                StartX = startX + startA,
                                StartY = startY + startB,
                                EndX = x + startA,
                                EndY = y + startB,
                                LastEdit = lastEdit
                            };
                        }
                    }
                }

                // reverse D-path
                for (var k = -D; k <= D; k += 2)
                {
                    var kIndex = k + HALF;
                    int x, y;
                    if (k == -D || (k != D && reverseDiagonal[kIndex + 1] <= reverseDiagonal[kIndex - 1]))
                    {
                        x = reverseDiagonal[kIndex + 1] - 1; // move left from k+1 diagonal
                        lastEdit = Edit.DeleteLeft;
                    }
                    else
                    {
                        x = reverseDiagonal[kIndex - 1]; //move up from k-1 diagonal
                        lastEdit = Edit.InsertUp;
                    }
                    y = x - (k + delta);

                    var endX = x;
                    var endY = y;

                    while (x > 0 && y > 0 && A[startA + x - 1] == B[startB + y - 1])
                    {
                        --x;
                        --y;
                    }

                    reverseDiagonal[kIndex] = x;

                    if (deltaEven && k + delta >= -D && k + delta <= D)
                    {
                        var forKIndex = (k + delta) + HALF;
                        var forX = forwardDiagonal[forKIndex];
                        var forY = forX - (k + delta);
                        if (forX >= x && forY >= y)
                        {
                            return new EditLengthResult
                            {
                                EditLength = 2 * D,
                                StartX = x + startA,
                                StartY = y + startB,
                                EndX = endX + startA,
                                EndY = endY + startB,
                                LastEdit = lastEdit
                            };
                        }
                    }
                }
            }

            throw new Exception("Should never get here");
        }

        private static string[] SmartSplit(string str, char[] delims)
        {
            var list = new List<string>();
            var begin = 0;
            for (var i = 0; i < str.Length; i++)
            {
                if (Array.IndexOf(delims, str[i]) != -1)
                {
                    list.Add(str.Substring(begin, i - begin));
                    list.Add(str.Substring(i, 1));
                    begin = i + 1;
                }
                else if (i >= str.Length - 1)
                {
                    list.Add(str.Substring(begin, i + 1 - begin));
                }
            }

            return list.ToArray();
        }

        private static string UppercaseVbaSymbols(string s, ref bool startInComment)
        {
            var sb = new StringBuilder(s.Length);
            var ls = startInComment ? LineState.Comment : LineState.New;
            startInComment = false;
            for (var i = 0; i < s.Length; i++)
            {
                if (ls == LineState.Quoted || ls == LineState.Comment)
                {
                    sb.Append(s[i]);
                }
                else
                {
                    sb.Append(char.ToUpper(s[i]));
                }
                switch (s[i])
                {
                case '"':
                    if (ls != LineState.Comment)
                    {
                        ls = ls == LineState.Quoted ? LineState.Normal : LineState.Quoted;
                    }
                    break;
                case '\'':
                    if (ls != LineState.Quoted && ls != LineState.Comment)
                    {
                        ls = LineState.Comment;
                    }
                    break;
                case ':':
                    if (ls != LineState.Quoted && ls != LineState.Comment)
                    {
                        ls = LineState.New;
                    }
                    break;
                case 'R':
                case 'r':
                    if (ls == LineState.New)
                    {
                        ls = LineState.CommentAwaitE;
                    }
                    else if (ls != LineState.Quoted && ls != LineState.Comment)
                    {
                        ls = LineState.Normal;
                    }
                    break;
                case 'E':
                case 'e':
                    if (ls == LineState.CommentAwaitE)
                    {
                        ls = LineState.CommentAwaitM;
                    }
                    else if (ls != LineState.Quoted && ls != LineState.Comment)
                    {
                        ls = LineState.Normal;
                    }
                    break;
                case 'M':
                case 'm':
                    if (ls == LineState.CommentAwaitM)
                    {
                        ls = LineState.CommentAwaitSpace;
                    }
                    else if (ls != LineState.Quoted && ls != LineState.Comment)
                    {
                        ls = LineState.Normal;
                    }
                    break;
                case ' ':
                    if (ls == LineState.CommentAwaitSpace)
                    {
                        ls = LineState.Comment;
                    }
                    else if (ls != LineState.Quoted && ls != LineState.Comment && ls != LineState.New)
                    {
                        ls = LineState.Normal;
                    }
                    break;
                case '_':
                    if (ls == LineState.Comment && i > 0 && i == s.Length - 1 && s[i - 1] == ' ')
                    {
                        startInComment = true;
                    }
                    else if (ls != LineState.Quoted && ls != LineState.Comment)
                    {
                        ls = LineState.Normal;
                    }
                    break;
                default:
                    if (ls != LineState.Quoted && ls != LineState.Comment)
                    {
                        ls = LineState.Normal;
                    }
                    break;
                }
            }
            return sb.ToString();
        }
    }

    internal class DiffBlock
    {
        public DiffBlock(int deleteStartA, int deleteCountA, int insertStartB, int insertCountB)
        {
            DeleteStartA = deleteStartA;
            DeleteCountA = deleteCountA;
            InsertStartB = insertStartB;
            InsertCountB = insertCountB;
        }

        public int DeleteCountA { get; }
        public int DeleteStartA { get; }
        public int InsertCountB { get; }
        public int InsertStartB { get; }
    }

    internal class DiffResult
    {
        public DiffResult(string[] peicesOld, string[] piecesNew, IList<DiffBlock> blocks)
        {
            PiecesOld = peicesOld;
            PiecesNew = piecesNew;
            DiffBlocks = blocks;
        }

        public IList<DiffBlock> DiffBlocks { get; }
        public string[] PiecesNew { get; }
        public string[] PiecesOld { get; }
    }

    internal class EditLengthResult
    {
        public int EditLength { get; set; }
        public int EndX { get; set; }
        public int EndY { get; set; }
        public Edit LastEdit { get; set; }
        public int StartX { get; set; }
        public int StartY { get; set; }
    }

    internal class ModificationData
    {
        public ModificationData(string str)
        {
            RawData = str;
        }

        public int[] HashedPieces { get; set; }
        public bool[] Modifications { get; set; }
        public string[] Pieces { get; set; }
        public string RawData { get; }
    }
}
